<#
.SYNOPSIS
    Importiert zuvor exportierte Microsoft Planner-Daten zurück in neue Pläne.

.DESCRIPTION
    Liest die JSON-Exportdateien von Export-PlannerData.ps1 und erstellt neue Pläne,
    Buckets, Tasks (inkl. Checklisten, Beschreibungen, Zuweisungen und Labels)
    in den angegebenen Microsoft 365 Gruppen.

.PARAMETER ImportPath
    Pfad zum Export-Verzeichnis (enthält die JSON-Dateien und _ExportIndex.json).
    Wird kein Pfad angegeben, werden Standard-Speicherorte automatisch durchsucht.

.PARAMETER TargetGroupId
    Gruppen-ID der Ziel-M365-Gruppe. Wird diese angegeben, werden alle Pläne
    in diese Gruppe importiert (statt in die Original-Gruppen aus dem Export).

.PARAMETER UserMapping
    Hashtable zur Abbildung alter Benutzer-IDs auf neue (für Tenant-Migrationen).
    Beispiel: @{ "alte-id" = "neue-id" }

.PARAMETER SkipAssignments
    Benutzerzuweisungen werden beim Import nicht wiederhergestellt.

.PARAMETER SkipCompletedTasks
    Abgeschlossene Tasks (percentComplete = 100) werden nicht importiert.

.PARAMETER DryRun
    Vorschau-Modus: Es werden keine Änderungen vorgenommen.
    Ohne -ImportPath: Listet alle gefundenen Exporte mit Gruppen und Tasks auf.
    Mit -ImportPath: Validiert den Export und zeigt was importiert werden würde.

.PARAMETER ListGroups
    Listet alle M365-Gruppen im Tenant mit ID und E-Mail auf und beendet das Script.
    Nützlich um die richtige -TargetGroupId zu ermitteln.

.PARAMETER TenantId
    Azure AD Tenant-ID für die Authentifizierung. Wird beim ersten erfolgreichen
    Login automatisch in ~/.planner-auth.json gespeichert und bei späteren
    Ausführungen für lautlose Anmeldung (ohne Browser) wiederverwendet.

.PARAMETER ThrottleDelayMs
    Verzögerung in Millisekunden zwischen API-Requests (Standard: 500).
    Bei 429-Fehlern wird automatisch der Retry-After-Wert verwendet.

.EXAMPLE
    .\Import-PlannerData.ps1
    Sucht automatisch nach Exporten und zeigt interaktive Auswahl.

.EXAMPLE
    .\Import-PlannerData.ps1 -DryRun
    Listet alle gefundenen Exporte mit Gruppen/Tasks auf (kein Browser nötig).

.EXAMPLE
    .\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260304_091500"
    Importiert alle Pläne aus dem angegebenen Export

.EXAMPLE
    .\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260304_091500" -DryRun
    Vorschau: Zeigt was importiert werden würde, ohne Änderungen vorzunehmen.

.EXAMPLE
    .\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260304_091500" -TargetGroupId "abc-123"
    Importiert alle Pläne in eine bestimmte Zielgruppe.

.EXAMPLE
    .\Import-PlannerData.ps1 -ListGroups
    Listet alle M365-Gruppen im Tenant auf.

.NOTES
    Voraussetzungen:
    - PowerShell 5.1 oder höher (empfohlen: PowerShell 7+)
    - Microsoft.Graph PowerShell Module
    - Berechtigungen: Group.ReadWrite.All, Tasks.ReadWrite

    WICHTIG:
    - Planner hat API-Rate-Limits. Das Script wartet automatisch zwischen Requests.
    - Anhänge/Referenzen (URLs) werden als Links wiederhergestellt.
    - Dateianhänge aus SharePoint müssen manuell neu verknüpft werden.
    - Benutzer-Zuweisungen funktionieren nur, wenn die User-IDs in der neuen
      Umgebung identisch sind (gleicher Tenant) oder ein Mapping bereitgestellt wird.

.AUTHOR
    Alexander Waller
    Datum: 2026-02-09
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ImportPath,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ([string]::IsNullOrEmpty($_)) {
            return $true  # Allow empty/null
        }
        if ($_ -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            throw "TargetGroupId muss eine gültige GUID sein.`n`nBeispiel: 12345678-1234-1234-1234-123456789abc`n`nSie haben eingegeben: $_`n`nUm die Gruppen-ID zu finden, verwenden Sie:`n  Get-MgGroup -Filter `"displayName eq 'Ihr Gruppenname'`" | Select-Object Id, DisplayName"
        }
        return $true
    })]
    [string]$TargetGroupId,

    [Parameter(Mandatory = $false)]
    [hashtable]$UserMapping,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAssignments,

    [Parameter(Mandatory = $false)]
    [switch]$SkipCompletedTasks,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$ListGroups,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$ThrottleDelayMs = 500
)

# Script-level variables für Error Tracking und Caching
$script:errorTracker = @{
    Plans = @{ Attempted = 0; Succeeded = 0; Failed = @() }
    Buckets = @{ Attempted = 0; Succeeded = 0; Failed = @() }
    Tasks = @{ Attempted = 0; Succeeded = 0; Failed = @() }
    TaskDetails = @{ Attempted = 0; Succeeded = 0; Failed = @() }
    UserResolution = @{ Attempted = 0; CacheHits = 0; Succeeded = 0; Failed = @() }
    Categories = @{
        NetworkErrors = @()
        PermissionErrors = @()
        DataValidationErrors = @()
        UnknownErrors = @()
    }
}

$script:userResolveCache = @{}
$script:groupPlansCache = @{}         # Cache: GroupId -> Array bestehender Pläne
$script:existingPlanDefaultAction = $null  # 'u'=überspringen, 'ue'=überschreiben, 'v'=verwenden (für alle)

#region Funktionen

function Write-PlannerLog {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $(switch ($Level) {
        "ERROR"  { "Red" }
        "WARN"   { "Yellow" }
        "OK"     { "Green" }
        "DRYRUN" { "Magenta" }
        default  { "White" }
    })
}

function Test-SafePath {
    <#
    .SYNOPSIS
        Validiert einen Dateisystempfad auf Sicherheit und Zugänglichkeit
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Export', 'Import')]
        [string]$Mode = 'Export',

        [Parameter(Mandatory = $false)]
        [switch]$AllowCreate,

        [Parameter(Mandatory = $false)]
        [ref]$ErrorMessage
    )

    # 1. Null/Leer-Check
    if ([string]::IsNullOrWhiteSpace($Path)) {
        if ($ErrorMessage) { $ErrorMessage.Value = "Pfad darf nicht leer sein" }
        return $false
    }

    # 2. UNC-Pfad blockieren (Sicherheit)
    if ($Path -match '^\\\\') {
        if ($ErrorMessage) { $ErrorMessage.Value = "UNC-Pfade (Netzwerkpfade) sind aus Sicherheitsgründen nicht erlaubt: $Path" }
        return $false
    }

    # 3. Pfad normalisieren (Relative Pfade auflösen, .. entfernen)
    try {
        # Relative Pfade relativ zum PowerShell-Arbeitsverzeichnis auflösen
        if (-not [System.IO.Path]::IsPathRooted($Path)) {
            $Path = Join-Path (Get-Location).Path $Path
        }
        $normalizedPath = [System.IO.Path]::GetFullPath($Path)
    }
    catch {
        if ($ErrorMessage) { $ErrorMessage.Value = "Ungültiges Pfad-Format: $($_.Exception.Message)" }
        return $false
    }

    # 4. Modus-spezifische Validierung
    if ($Mode -eq 'Export') {
        # Export: Pfad muss schreibbar sein oder erstellt werden können
        if (Test-Path $normalizedPath) {
            # Existiert bereits - muss Verzeichnis sein
            if (-not (Test-Path $normalizedPath -PathType Container)) {
                if ($ErrorMessage) { $ErrorMessage.Value = "Pfad existiert bereits als Datei (kein Verzeichnis): $normalizedPath" }
                return $false
            }

            # Schreibrechte testen
            try {
                $testFile = Join-Path $normalizedPath ".write_test_$([guid]::NewGuid().ToString('N').Substring(0,8))"
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            }
            catch {
                if ($ErrorMessage) { $ErrorMessage.Value = "Keine Schreibrechte für Verzeichnis: $normalizedPath" }
                return $false
            }
        }
        else {
            # Existiert nicht - übergeordnetes Verzeichnis prüfen
            $parentPath = Split-Path $normalizedPath -Parent

            if (-not $parentPath) {
                if ($ErrorMessage) { $ErrorMessage.Value = "Kann übergeordnetes Verzeichnis nicht ermitteln für: $normalizedPath" }
                return $false
            }

            if (-not (Test-Path $parentPath)) {
                if ($AllowCreate) {
                    # Prüfe ob Großeltern-Verzeichnis existiert (max 1 Ebene erstellen)
                    $grandparentPath = Split-Path $parentPath -Parent
                    if ($grandparentPath -and -not (Test-Path $grandparentPath)) {
                        if ($ErrorMessage) { $ErrorMessage.Value = "Übergeordnetes Verzeichnis existiert nicht: $grandparentPath (maximal 1 Ebene kann automatisch erstellt werden)" }
                        return $false
                    }
                }
                else {
                    if ($ErrorMessage) { $ErrorMessage.Value = "Übergeordnetes Verzeichnis existiert nicht: $parentPath" }
                    return $false
                }
            }

            # Schreibrechte für übergeordnetes Verzeichnis testen
            try {
                $testParentPath = if (Test-Path $parentPath) { $parentPath } else { Split-Path $parentPath -Parent }
                $testFile = Join-Path $testParentPath ".write_test_$([guid]::NewGuid().ToString('N').Substring(0,8))"
                [System.IO.File]::WriteAllText($testFile, "test")
                Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            }
            catch {
                if ($ErrorMessage) { $ErrorMessage.Value = "Keine Schreibrechte für übergeordnetes Verzeichnis: $testParentPath" }
                return $false
            }
        }
    }
    elseif ($Mode -eq 'Import') {
        # Import: Pfad muss existieren und lesbar sein
        if (-not (Test-Path $normalizedPath)) {
            if ($ErrorMessage) { $ErrorMessage.Value = "Import-Verzeichnis existiert nicht: $normalizedPath" }
            return $false
        }

        # Muss Verzeichnis sein
        if (-not (Test-Path $normalizedPath -PathType Container)) {
            if ($ErrorMessage) { $ErrorMessage.Value = "Import-Pfad ist kein Verzeichnis: $normalizedPath" }
            return $false
        }

        # Leserechte testen
        try {
            Get-ChildItem $normalizedPath -ErrorAction Stop | Out-Null
        }
        catch {
            if ($ErrorMessage) { $ErrorMessage.Value = "Keine Leserechte für Import-Verzeichnis: $normalizedPath" }
            return $false
        }
    }

    return $true
}

function Add-ErrorToTracker {
    param(
        [ValidateSet('Plan', 'Bucket', 'Task', 'TaskDetail', 'UserResolution')]
        [string]$ItemType,
        [string]$ItemName,
        [object]$Exception,
        [string]$Context
    )

    $errorDetails = @{
        ItemType = $ItemType
        ItemName = $ItemName
        Context = $Context
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Message = if ($Exception) {
            # Versuche den detaillierten Fehlertext zu extrahieren
            $msg = ''
            if ($Exception.PSObject.Properties['Exception'] -and $Exception.Exception) {
                $msg = $Exception.Exception.Message
            }
            if (-not $msg) { $msg = $Exception.Message }
            # Aus JSON-Body den 'message'-Wert extrahieren wenn vorhanden
            if ($msg -match '"message"\s*:\s*"([^"]+)"') { $msg = $Matches[1] }
            if ($msg) { $msg } else { "Unbekannter Fehler" }
        } else { "Unbekannter Fehler" }
        ExceptionType = if ($Exception) { $Exception.GetType().FullName } else { "N/A" }
    }

    # Add StatusCode if available
    if ($null -ne $Exception -and $null -ne $Exception.Exception -and
        $null -ne $Exception.Exception.Response) {
        try {
            $errorDetails.StatusCode = $Exception.Exception.Response.StatusCode.value__
        }
        catch {
            $errorDetails.StatusCode = "N/A"
        }
    }
    else {
        $errorDetails.StatusCode = "N/A"
    }

    # Add to specific item type failures
    $itemTypePlural = if ($ItemType -eq 'UserResolution') { 'UserResolution' } else { "$ItemType" + "s" }
    $script:errorTracker.$itemTypePlural.Failed += $errorDetails

    # Categorize error
    $errorMessage = if ($errorDetails.Message) { $errorDetails.Message.ToLower() } else { '' }
    $statusCode = $errorDetails.StatusCode

    if ($statusCode -in @(408, 500, 502, 503, 504, "N/A") -or
        $errorMessage -match "(timeout|timed out|connection|network|dns)") {
        $script:errorTracker.Categories.NetworkErrors += $errorDetails
    }
    elseif ($statusCode -in @(401, 403) -or
            $errorMessage -match "(unauthorized|forbidden|permission|access denied)") {
        $script:errorTracker.Categories.PermissionErrors += $errorDetails
    }
    elseif ($statusCode -in @(400, 422) -or
            $errorMessage -match "(validation|invalid|bad request|malformed)") {
        $script:errorTracker.Categories.DataValidationErrors += $errorDetails
    }
    else {
        $script:errorTracker.Categories.UnknownErrors += $errorDetails
    }
}

function Write-ErrorSummary {
    param([string]$OutputPath)

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  FEHLER-ZUSAMMENFASSUNG" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""

    # Overall statistics
    $totalFailed = $script:errorTracker.Plans.Failed.Count +
                   $script:errorTracker.Buckets.Failed.Count +
                   $script:errorTracker.Tasks.Failed.Count

    Write-Host "Gesamtstatistik:" -ForegroundColor Yellow
    Write-Host "  Pläne:    $($script:errorTracker.Plans.Succeeded)/$($script:errorTracker.Plans.Attempted) erfolgreich" -ForegroundColor $(if ($script:errorTracker.Plans.Failed.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Buckets:  $($script:errorTracker.Buckets.Succeeded)/$($script:errorTracker.Buckets.Attempted) erfolgreich" -ForegroundColor $(if ($script:errorTracker.Buckets.Failed.Count -eq 0) { "Green" } else { "Red" })
    Write-Host "  Tasks:    $($script:errorTracker.Tasks.Succeeded)/$($script:errorTracker.Tasks.Attempted) erfolgreich" -ForegroundColor $(if ($script:errorTracker.Tasks.Failed.Count -eq 0) { "Green" } else { "Red" })

    if ($script:errorTracker.UserResolution.Attempted -gt 0) {
        Write-Host "  Benutzer: $($script:errorTracker.UserResolution.Succeeded)/$($script:errorTracker.UserResolution.Attempted) aufgelöst (Cache-Hits: $($script:errorTracker.UserResolution.CacheHits))" -ForegroundColor $(if ($script:errorTracker.UserResolution.Failed.Count -eq 0) { "Green" } else { "Yellow" })
    }
    Write-Host ""

    # Error breakdown by category
    if ($totalFailed -gt 0) {
        Write-Host "Fehler nach Kategorie:" -ForegroundColor Yellow
        if ($script:errorTracker.Categories.NetworkErrors.Count -gt 0) {
            Write-Host "  Netzwerkfehler:         $($script:errorTracker.Categories.NetworkErrors.Count)" -ForegroundColor Red
        }
        if ($script:errorTracker.Categories.PermissionErrors.Count -gt 0) {
            Write-Host "  Berechtigungsfehler:    $($script:errorTracker.Categories.PermissionErrors.Count)" -ForegroundColor Red
        }
        if ($script:errorTracker.Categories.DataValidationErrors.Count -gt 0) {
            Write-Host "  Validierungsfehler:     $($script:errorTracker.Categories.DataValidationErrors.Count)" -ForegroundColor Red
        }
        if ($script:errorTracker.Categories.UnknownErrors.Count -gt 0) {
            Write-Host "  Unbekannte Fehler:      $($script:errorTracker.Categories.UnknownErrors.Count)" -ForegroundColor Red
        }
        Write-Host ""

        # Detailed failure list (limit to first 10 per type)
        Write-Host "Fehlgeschlagene Elemente (max. 10 pro Typ):" -ForegroundColor Yellow

        $displayedPlans = 0
        foreach ($failure in $script:errorTracker.Plans.Failed) {
            if ($displayedPlans++ -ge 10) { break }
            Write-Host "  [PLAN] $($failure.ItemName)" -ForegroundColor Red
            Write-Host "    Fehler: $($failure.Message)" -ForegroundColor Gray
        }

        $displayedBuckets = 0
        foreach ($failure in $script:errorTracker.Buckets.Failed) {
            if ($displayedBuckets++ -ge 10) { break }
            Write-Host "  [BUCKET] $($failure.ItemName)" -ForegroundColor Red
            Write-Host "    Kontext: $($failure.Context)" -ForegroundColor Gray
            Write-Host "    Fehler: $($failure.Message)" -ForegroundColor Gray
        }

        $displayedTasks = 0
        foreach ($failure in $script:errorTracker.Tasks.Failed) {
            if ($displayedTasks++ -ge 10) { break }
            Write-Host "  [TASK] $($failure.ItemName)" -ForegroundColor Red
            Write-Host "    Kontext: $($failure.Context)" -ForegroundColor Gray
            Write-Host "    Fehler: $($failure.Message)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    else {
        Write-Host "Keine Fehler aufgetreten!" -ForegroundColor Green
        Write-Host ""
    }

    # Write JSON error report
    $errorReport = @{
        Timestamp = Get-Date -Format "o"
        Summary = @{
            TotalAttempted = $script:errorTracker.Plans.Attempted + $script:errorTracker.Buckets.Attempted + $script:errorTracker.Tasks.Attempted
            TotalSucceeded = $script:errorTracker.Plans.Succeeded + $script:errorTracker.Buckets.Succeeded + $script:errorTracker.Tasks.Succeeded
            TotalFailed = $totalFailed
        }
        Details = $script:errorTracker
    }

    $reportPath = Join-Path $OutputPath "import_errors.json"
    try {
        $errorReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding utf8 -ErrorAction Stop
        Write-Host "Detaillierter Fehlerbericht gespeichert: $reportPath" -ForegroundColor Cyan
    }
    catch {
        Write-PlannerLog "Warnung: Konnte Fehlerbericht nicht speichern: $_" "WARN"
    }

    # Determine exit code
    if ($script:errorTracker.Plans.Failed.Count -gt 0) {
        return 2  # Total failure
    }
    elseif ($totalFailed -gt 0) {
        return 1  # Partial failure
    }
    else {
        return 0  # Success
    }
}

function Write-CacheStatistics {
    $totalLookups = $script:errorTracker.UserResolution.Attempted
    $cacheHits = $script:errorTracker.UserResolution.CacheHits
    $cacheMisses = $totalLookups - $cacheHits

    if ($totalLookups -gt 0) {
        $hitRate = [math]::Round(($cacheHits / $totalLookups) * 100, 2)
        $apiCallsSaved = $cacheHits * 2

        Write-PlannerLog "Benutzer-Auflösung Cache-Statistik:" "OK"
        Write-PlannerLog "  Gesamt-Lookups:     $totalLookups" "INFO"
        Write-PlannerLog "  Cache-Hits:         $cacheHits" "INFO"
        Write-PlannerLog "  Cache-Misses:       $cacheMisses" "INFO"
        Write-PlannerLog "  Trefferquote:       $hitRate%" "INFO"
        Write-PlannerLog "  Eingesparte API-Calls (geschätzt): $apiCallsSaved" "OK"
    }
}

function ConvertTo-IsoDate {
    <#
    .SYNOPSIS
        Konvertiert ein Datum in ISO 8601 Format für die Graph API.
        Unterstützt DateTime-Objekte (ConvertFrom-Json konvertiert /Date(ms)/ automatisch),
        /Date(ms)/-Strings (OData v3) und ISO-Strings.
    #>
    param([object]$Value)
    if (-not $Value) { return $null }
    # ConvertFrom-Json wandelt /Date(ms)/ bereits in DateTime/DateTimeOffset um
    if ($Value -is [datetime]) {
        return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    if ($Value -is [datetimeoffset]) {
        return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    $str = $Value.ToString().Trim()
    # OData v3 Fallback (falls doch als String): /Date(1769554800000)/
    if ($str -match '^/Date\((\d+)') {
        $ms = [long]$Matches[1]
        return [DateTimeOffset]::FromUnixTimeMilliseconds($ms).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    # ISO-String Fallback
    try {
        return ([datetime]::Parse($str, [System.Globalization.CultureInfo]::InvariantCulture)).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
    catch {
        return $null
    }
}

function Get-OrCreateWarningCategory {
    <#
    .SYNOPSIS
        Findet einen freien Kategorie-Slot im Plan und legt eine "⚠ Import-Fehler"-Kategorie an.
        Gibt den Kategorie-Schlüssel zurück (z.B. "category25"), oder $null bei Fehler.
    #>
    param(
        [string]$PlanId,
        [object]$ExistingCategories
    )
    # Bereits verwendete Slots bestimmen
    $usedSlots = @{}
    if ($ExistingCategories) {
        $ExistingCategories.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object {
            $usedSlots[$_.Name] = $true
        }
    }

    # Freien Slot von hinten suchen (category25..category1)
    $freeSlot = $null
    for ($i = 25; $i -ge 1; $i--) {
        $key = "category$i"
        if (-not $usedSlots.ContainsKey($key)) {
            $freeSlot = $key
            break
        }
    }

    if (-not $freeSlot) {
        Write-PlannerLog "  Kein freier Kategorie-Slot für Warn-Label verfügbar (alle 25 belegt)" "WARN"
        return $null
    }

    try {
        $planDetails = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/details"
        
        # Hashtable explizit aufbauen (Variable als Key)
        $bodyObj = @{ categoryDescriptions = @{} }
        $bodyObj.categoryDescriptions[$freeSlot] = "⚠ Import-Fehler"
        
        # JSON-Body erstellen
        $jsonBody = $bodyObj | ConvertTo-Json -Depth 5 -Compress
        
        $patchParams = @{
            Method      = "PATCH"
            Uri         = "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/details"
            Body        = $jsonBody
            ContentType = "application/json"
            Headers     = @{ "If-Match" = $planDetails.'@odata.etag' }
            OutputType  = "PSObject"
        }
        Invoke-MgGraphRequest @patchParams
        Write-PlannerLog "  Warn-Label erstellt in Slot '$freeSlot'" "OK"
        return $freeSlot
    }
    catch {
        Write-PlannerLog "  [DEBUG] Warn-Label Erstellung Fehler Details:" "WARN"
        Write-PlannerLog "    - FreeSlot: '$freeSlot' (Typ: $($freeSlot.GetType().Name))" "WARN"
        Write-PlannerLog "    - BodyObj.categoryDescriptions Keys: $($bodyObj.categoryDescriptions.Keys -join ', ')" "WARN"
        Write-PlannerLog "    - BodyObj Typ: $($bodyObj.GetType().Name)" "WARN"
        Write-PlannerLog "    - BodyObj.categoryDescriptions Typ: $($bodyObj.categoryDescriptions.GetType().Name)" "WARN"
        Write-PlannerLog "    - JSON Body: $jsonBody" "WARN"
        Write-PlannerLog "    - Fehlermeldung: $_" "WARN"
        Write-PlannerLog "    - Exception Typ: $($_.Exception.GetType().FullName)" "WARN"
        if ($_.Exception.InnerException) {
            Write-PlannerLog "    - Inner Exception: $($_.Exception.InnerException.Message)" "WARN"
        }
        return $null
    }
}

function Get-ExistingPlansForGroup {
    <#
    .SYNOPSIS
        Ruft alle vorhandenen Planner-Pläne einer M365-Gruppe ab (mit Caching).
    #>
    param([string]$GroupId)
    if ($script:groupPlansCache.ContainsKey($GroupId)) {
        return $script:groupPlansCache[$GroupId]
    }
    try {
        $response = Invoke-GraphWithRetry -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/planner/plans"
        $plans = if ($response -and $response.value) { @($response.value) } else { @() }
        $script:groupPlansCache[$GroupId] = $plans
        return $plans
    }
    catch {
        Write-PlannerLog "Konnte bestehende Pläne für Gruppe $GroupId nicht abrufen: $_" "WARN"
        return @()
    }
}

function Invoke-GraphWithRetry {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [hashtable]$Headers,
        [int]$MaxRetries = 3
    )

    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            Start-Sleep -Milliseconds $ThrottleDelayMs

            $params = @{
                Method     = $Method
                Uri        = $Uri
                OutputType = "PSObject"
            }

            if ($Body) {
                $params["Body"] = ($Body | ConvertTo-Json -Depth 20)
                $params["ContentType"] = "application/json"
            }

            if ($Headers) {
                $params["Headers"] = $Headers
            }

            return Invoke-MgGraphRequest @params
        }
        catch {
            $attempt++

            # Defensive null checks before accessing Response properties
            $statusCode = $null
            $hasResponse = $false

            if ($null -ne $_.Exception -and $null -ne $_.Exception.Response) {
                $hasResponse = $true
                try {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                catch {
                    Write-PlannerLog "Warnung: StatusCode nicht verfügbar: $($_.Exception.Message)" "WARN"
                }
            }

            # Log exception details for debugging
            $exceptionType = if ($_.Exception) { $_.Exception.GetType().FullName } else { "Unknown" }
            $errorMessage = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }

            # Detect transient vs permanent errors
            $isTransient = $false
            $isRateLimit = $false

            if ($statusCode -eq 429 -or $errorMessage -match "429") {
                $isRateLimit = $true
                $isTransient = $true
            }
            elseif ($statusCode -in @(408, 500, 502, 503, 504) -or
                    $errorMessage -match "(timeout|timed out|connection|network|temporary)") {
                $isTransient = $true
            }
            elseif ($statusCode -in @(400, 401, 403, 404) -or
                    $errorMessage -match "(unauthorized|forbidden|not found|bad request)") {
                $isTransient = $false
            }
            else {
                $isTransient = ($attempt -lt $MaxRetries)
            }

            # Handle rate limiting (429)
            if ($isRateLimit) {
                $retryAfter = 30

                if ($hasResponse -and $null -ne $_.Exception.Response.Headers) {
                    try {
                        $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                        if ($retryAfterHeader) {
                            $retryAfter = [int]$retryAfterHeader
                        }
                    }
                    catch {
                        Write-PlannerLog "Warnung: Retry-After Header konnte nicht gelesen werden, verwende Standard: $retryAfter Sekunden" "WARN"
                    }
                }

                Write-PlannerLog "Rate Limited (429). Warte $retryAfter Sekunden... (Versuch $attempt/$MaxRetries)" "WARN"
                Write-PlannerLog "  Fehlerdetails: $errorMessage" "WARN"
                Start-Sleep -Seconds $retryAfter
            }
            elseif ($attempt -ge $MaxRetries) {
                Write-PlannerLog "Maximale Anzahl an Wiederholungen erreicht ($MaxRetries). Fehlertyp: $exceptionType" "ERROR"
                Write-PlannerLog "  URI: $Uri" "ERROR"
                Write-PlannerLog "  StatusCode: $(if ($statusCode) { $statusCode } else { 'N/A' })" "ERROR"
                Write-PlannerLog "  Nachricht: $errorMessage" "ERROR"
                throw $_
            }
            elseif ($isTransient) {
                $waitSeconds = 2 * $attempt
                Write-PlannerLog "Vorübergehender Fehler bei Graph-Request (Versuch $attempt/$MaxRetries)" "WARN"
                Write-PlannerLog "  Fehlertyp: $exceptionType" "WARN"
                Write-PlannerLog "  StatusCode: $(if ($statusCode) { $statusCode } else { 'N/A' })" "WARN"
                Write-PlannerLog "  Nachricht: $errorMessage" "WARN"
                Write-PlannerLog "  Warte $waitSeconds Sekunden vor erneutem Versuch..." "WARN"
                Start-Sleep -Seconds $waitSeconds
            }
            else {
                Write-PlannerLog "Permanenter Fehler erkannt, breche ab" "ERROR"
                Write-PlannerLog "  Fehlertyp: $exceptionType" "ERROR"
                Write-PlannerLog "  StatusCode: $(if ($statusCode) { $statusCode } else { 'N/A' })" "ERROR"
                Write-PlannerLog "  Nachricht: $errorMessage" "ERROR"
                throw $_
            }
        }
    }
}

function Connect-ToGraph {
    $script:AuthConfigPath = Join-Path $env:USERPROFILE '.planner-auth.json'

    # Auth-Config laden/speichern
    function Get-PlannerAuthConfig {
        if (Test-Path $script:AuthConfigPath) {
            try { return Get-Content $script:AuthConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json }
            catch { return $null }
        }
        return $null
    }
    function Save-PlannerAuthConfig($tid, $account) {
        try {
            @{ TenantId = $tid; Account = $account; SavedAt = (Get-Date -Format 'o') } |
                ConvertTo-Json | Out-File $script:AuthConfigPath -Encoding UTF8 -Force
        } catch { }
    }

    Write-PlannerLog "Verbinde mit Microsoft Graph..."
    $scopes = "Group.ReadWrite.All", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All"

    try {
        # Prüfe ob bereits verbunden (gleiche Session)
        $context = Get-MgContext
        if ($null -ne $context -and -not [string]::IsNullOrEmpty($context.Account)) {
            Write-PlannerLog "Bereits verbunden als: $($context.Account)" "OK"
            return $true
        }

        # TenantId + Account: Parameter > Config-Datei
        $cfg = Get-PlannerAuthConfig
        $savedAccount = if ($cfg) { $cfg.Account } else { $null }

        $tid = if (-not [string]::IsNullOrEmpty($TenantId)) {
            $TenantId
        } else {
            if ($cfg -and $cfg.TenantId) {
                Write-PlannerLog "Verwende gespeicherte Anmeldedaten ($($cfg.Account))..."
                $cfg.TenantId
            } else { $null }
        }

        $connected = $false

        # 1. Versuch: Silent-Auth via MSAL-Cache (nur wenn TenantId bekannt)
        if ($tid) {
            try {
                $silentArgs = @{ TenantId = $tid; Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
                Connect-MgGraph @silentArgs
                $connected = $true
            } catch {
                $errMsg = $_.Exception.Message
                if ($_.Exception.InnerException) { $errMsg += ' ' + $_.Exception.InnerException.Message }
                $isCancelled = $errMsg -match 'authentication_canceled|user_cancelled|UserCanceled' `
                    -or $errMsg -match 'The user did not complete the authentication' `
                    -or $errMsg -match 'Authentication was canceled by the user' `
                    -or $errMsg -match 'abgebrochen|aborted|cancelled|canceled' `
                    -or $_ -is [System.OperationCanceledException] `
                    -or $_.Exception.InnerException -is [System.OperationCanceledException]
                if ($isCancelled) {
                    Write-Host ""
                    Write-Host "Anmeldung wurde abgebrochen. Script wird beendet." -ForegroundColor Red
                    Write-Host ""
                    exit 1
                }
                Write-PlannerLog "Silent-Auth fehlgeschlagen, Anmeldung erforderlich..." "WARN"
            }
        }

        # 2. Interaktive Anmeldung: Benutzer wählt Browser oder Device Code
        if (-not $connected) {
            Write-Host ""
            Write-Host "  Anmeldung erforderlich" -ForegroundColor Yellow
            if ($savedAccount) { Write-Host "  Konto: $savedAccount" -ForegroundColor Gray }
            Write-Host ""
            Write-Host "  [B] Browser-Anmeldung  (öffnet Browserfenster)" -ForegroundColor Cyan
            Write-Host "  [D] Device Code        (Code im Browser eingeben) [Standard]" -ForegroundColor Cyan
            Write-Host ""
            $authChoice = (Read-Host "  Auswahl [B/D]").Trim().ToUpper()
            if ([string]::IsNullOrEmpty($authChoice)) { $authChoice = 'D' }
            Write-Host ""

            try {
                if ($authChoice -eq 'B') {
                    $connectArgs = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
                    if ($tid) { $connectArgs['TenantId'] = $tid }
                    Connect-MgGraph @connectArgs
                } else {
                    $connectArgs = @{ Scopes = $scopes; UseDeviceCode = $true; NoWelcome = $true; ErrorAction = 'Stop' }
                    if ($tid) { $connectArgs['TenantId'] = $tid }
                    Connect-MgGraph @connectArgs
                }
                $connected = $true
            } catch {
                $errMsg = $_.Exception.Message
                if ($_.Exception.InnerException) { $errMsg += ' ' + $_.Exception.InnerException.Message }
                $isCancelled = $errMsg -match 'authentication_canceled|user_cancelled|UserCanceled' `
                    -or $errMsg -match 'The user did not complete the authentication' `
                    -or $errMsg -match 'Authentication was canceled by the user' `
                    -or $errMsg -match 'abgebrochen|aborted|cancelled|canceled' `
                    -or $_ -is [System.OperationCanceledException] `
                    -or $_.Exception.InnerException -is [System.OperationCanceledException]
                if ($isCancelled) {
                    Write-Host ""
                    Write-Host "Anmeldung wurde abgebrochen. Script wird beendet." -ForegroundColor Red
                    Write-Host ""
                    exit 1
                }
                throw
            }
        }

        $context = Get-MgContext
        if ($null -eq $context -or [string]::IsNullOrEmpty($context.Account)) {
            throw "Keine gültige Verbindung hergestellt"
        }

        # TenantId + Account für nächste Ausführung speichern
        $resolvedTid = if ($tid) { $tid } else { $context.TenantId }
        Save-PlannerAuthConfig -tid $resolvedTid -account $context.Account

        Write-PlannerLog "Verbunden als: $($context.Account)" "OK"
        return $true
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($_.Exception.InnerException) { $errMsg += ' ' + $_.Exception.InnerException.Message }
        $isCancelled = $errMsg -match 'authentication_canceled|user_cancelled|UserCanceled' `
            -or $errMsg -match 'The user did not complete the authentication' `
            -or $errMsg -match 'Authentication was canceled by the user' `
            -or $errMsg -match 'abgebrochen|aborted|cancelled|canceled' `
            -or $_ -is [System.OperationCanceledException] `
            -or $_.Exception.InnerException -is [System.OperationCanceledException]
        if ($isCancelled) {
            Write-Host ""
            Write-Host "Anmeldung wurde abgebrochen. Script wird beendet." -ForegroundColor Red
            Write-Host ""
            exit 1
        }
        Write-PlannerLog "Fehler bei der Verbindung: $_" "ERROR"
        return $false
    }
}

function Test-TargetGroup {
    param([string]$GroupId)

    if ([string]::IsNullOrEmpty($GroupId)) {
        return $true  # Keine Zielgruppe angegeben, wird später aus JSON ermittelt
    }

    Write-PlannerLog "Validiere Zielgruppe: $GroupId"
    try {
        $group = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId" -ErrorAction Stop
        Write-PlannerLog "  Gruppe gefunden: $($group.displayName)" "OK"
        Write-PlannerLog "  Gruppentyp: $($group.groupTypes -join ', ')" "OK"

        # Prüfe ob es eine M365-Gruppe ist (Unified)
        if ($group.groupTypes -notcontains "Unified") {
            Write-PlannerLog "  WARNUNG: Gruppe ist keine Microsoft 365 Gruppe. Planner benötigt M365-Gruppen." "WARN"
            return $false
        }

        return $true
    }
    catch {
        if ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*does not exist*" -or $_.Exception.Message -like "*not found*") {
            Write-PlannerLog "  Gruppe existiert nicht: $GroupId" "ERROR"
        }
        elseif ($_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*Forbidden*") {
            Write-PlannerLog "  Keine Berechtigung zum Zugriff auf Gruppe: $GroupId" "ERROR"
        }
        else {
            Write-PlannerLog "  Fehler beim Validieren der Gruppe: $_" "ERROR"
        }
        return $false
    }
}

function Resolve-UserId {
    param([string]$OldUserId, [hashtable]$OldUserMap)

    # Track resolution attempt
    $script:errorTracker.UserResolution.Attempted++

    # Check cache first
    if ($script:userResolveCache.ContainsKey($OldUserId)) {
        $script:errorTracker.UserResolution.CacheHits++
        $cached = $script:userResolveCache[$OldUserId]

        if ($cached.Status -eq "Success") {
            return $cached.NewUserId
        }
        else {
            # Previous lookup failed, don't retry
            return $null
        }
    }

    # UserMapping has highest priority
    if ($UserMapping -and $UserMapping.ContainsKey($OldUserId)) {
        $resolvedId = $UserMapping[$OldUserId]
        $script:userResolveCache[$OldUserId] = @{
            NewUserId = $resolvedId
            Status = "Success"
            Timestamp = Get-Date
            Method = "UserMapping"
        }
        $script:errorTracker.UserResolution.Succeeded++
        return $resolvedId
    }

    # Try to find user in new environment
    $resolvedId = $null

    if ($OldUserMap -and $OldUserMap[$OldUserId]) {
        $upn = $OldUserMap[$OldUserId].UserPrincipalName
        $mail = $OldUserMap[$OldUserId].Mail

        # Try UPN lookup
        if ($upn) {
            try {
                $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                if ($user -and $user.id) {
                    $resolvedId = $user.id
                }
            }
            catch {
                Write-PlannerLog "  Warnung: Benutzer konnte nicht per UPN gefunden werden: $upn" "WARN"
            }
        }

        # Try Mail lookup if UPN failed
        if (-not $resolvedId -and $mail -and $mail -ne $upn) {
            try {
                $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mail`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                if ($user -and $user.id) {
                    $resolvedId = $user.id
                }
            }
            catch {
                Write-PlannerLog "  Warnung: Benutzer konnte nicht per Mail gefunden werden: $mail" "WARN"
            }
        }
    }

    # Fallback: Try original ID
    if (-not $resolvedId) {
        try {
            $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$OldUserId`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
            if ($user -and $user.id) {
                $resolvedId = $user.id
            }
        }
        catch {
            Write-PlannerLog "  Warnung: Benutzer konnte nicht per ID gefunden werden: $OldUserId" "WARN"
        }
    }

    # Cache the result (success or failure)
    if ($resolvedId) {
        $script:userResolveCache[$OldUserId] = @{
            NewUserId = $resolvedId
            Status = "Success"
            Timestamp = Get-Date
        }
        $script:errorTracker.UserResolution.Succeeded++
        return $resolvedId
    }
    else {
        # Cache the failure to avoid repeated lookups
        $script:userResolveCache[$OldUserId] = @{
            NewUserId = $null
            Status = "Failed"
            Timestamp = Get-Date
        }

        # Add to error tracker
        $userName = if ($OldUserMap -and $OldUserMap[$OldUserId]) {
            $OldUserMap[$OldUserId].DisplayName
        } else {
            $OldUserId
        }
        Add-ErrorToTracker -ItemType "UserResolution" -ItemName $userName -Exception $null -Context "Benutzer-ID: $OldUserId"

        return $null
    }
}

function Import-PlanFromJson {
    param(
        [string]$JsonFilePath,
        [string]$TargetGroupId
    )

    Write-PlannerLog "Lade Export-Datei: $JsonFilePath"
    $planData = Get-Content $JsonFilePath -Raw -Encoding UTF8 | ConvertFrom-Json

    $planTitle = $planData.Plan.title
    $originalGroupId = $planData.Plan.groupId
    
    # Liste für fehlende Benutzerzuweisungen
    $missingUsers = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Zielgruppe bestimmen
    $groupId = if ($TargetGroupId) { $TargetGroupId } else { $originalGroupId }

    if (-not $groupId) {
        Write-PlannerLog "Keine Zielgruppe angegeben und keine Original-Gruppe gefunden!" "ERROR"
        return $null
    }

    # Validiere Gruppen-ID Format (originalGroupId stammt aus JSON, nicht vertrauenswürdig)
    if (-not $TargetGroupId -and $originalGroupId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
        Write-PlannerLog "Ungültige Gruppen-ID in Export-Datei: $originalGroupId" "ERROR"
        return $null
    }

    Write-PlannerLog "Erstelle Plan '$planTitle' in Gruppe $groupId..."

    if ($DryRun) {
        # Validiere Zielgruppe im DryRun
        if (-not (Test-TargetGroup -GroupId $groupId)) {
            Write-PlannerLog "[DRY RUN] FEHLER: Zielgruppe $groupId ist ungültig" "ERROR"
            return $null
        }

        Write-PlannerLog "[DRY RUN] Würde Plan '$planTitle' erstellen" "DRYRUN"
        Write-PlannerLog "[DRY RUN] Zielgruppe: $groupId" "DRYRUN"
        Write-PlannerLog "[DRY RUN] Buckets: $($planData.Buckets.Count)" "DRYRUN"
        Write-PlannerLog "[DRY RUN] Tasks: $($planData.Tasks.Count)" "DRYRUN"

        if (-not $SkipCompletedTasks) {
            $completedTasks = ($planData.Tasks | Where-Object { $_.percentComplete -eq 100 }).Count
            if ($completedTasks -gt 0) {
                Write-PlannerLog "[DRY RUN]   davon abgeschlossen: $completedTasks" "DRYRUN"
            }
        }

        if (-not $SkipAssignments) {
            $tasksWithAssignments = ($planData.Tasks | Where-Object { $_.assignments -and $_.assignments.PSObject.Properties.Count -gt 0 }).Count
            if ($tasksWithAssignments -gt 0) {
                Write-PlannerLog "[DRY RUN]   Tasks mit Zuweisungen: $tasksWithAssignments" "DRYRUN"
            }
        }

        return @{
            Status         = "DryRun"
            PlanTitle      = $planTitle
            GroupName      = $planData.Plan.groupDisplayName
            TargetGroupId  = $groupId
            BucketsCreated = @($planData.Buckets).Count
            TasksCreated   = @($planData.Tasks).Count
        }
    }

    # 1. Plan erstellen (oder vorhandenen verwenden/überschreiben)
    $script:errorTracker.Plans.Attempted++

    # Prüfen ob Plan mit gleichem Titel bereits in Zielgruppe existiert
    $existingPlan = Get-ExistingPlansForGroup -GroupId $groupId |
        Where-Object { $_.title -eq $planTitle } | Select-Object -First 1

    $newPlan = $null
    if ($existingPlan) {
        Write-PlannerLog "  Plan '$planTitle' existiert bereits (ID: $($existingPlan.id))" "WARN"

        # Aktion bestimmen: aus Cache oder interaktiv fragen
        $action = $script:existingPlanDefaultAction
        if (-not $action) {
            Write-Host ""
            Write-Host "  ╔══════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
            Write-Host "  ║  Plan bereits vorhanden: '$planTitle'" -ForegroundColor Yellow
            Write-Host "  ╚══════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
            Write-Host "  Vorhandener Plan-ID: $($existingPlan.id)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  [v] Verwenden      - Buckets/Tasks in vorhandenen Plan importieren" -ForegroundColor Cyan
            Write-Host "  [ue] Überschreiben - Vorhandenen Plan LÖSCHEN und neu erstellen" -ForegroundColor Red
            Write-Host "  [u] Überspringen   - Diesen Plan nicht importieren" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  Für alle weiteren Konflikte: Eingabe + 'a' (z.B. 'va', 'uea', 'ua')" -ForegroundColor DarkGray
            Write-Host ""
            $choice = (Read-Host "  Auswahl").Trim().ToLower()

            if ($choice -match 'a$') {
                $script:existingPlanDefaultAction = $choice -replace 'a$', ''
            }
            $action = $choice -replace 'a$', ''
        }

        switch ($action) {
            'ue' {
                Write-PlannerLog "  Lösche vorhandenen Plan '$planTitle'..." "INFO"
                try {
                    $planToDelete = Invoke-GraphWithRetry -Method GET `
                        -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($existingPlan.id)"
                    Invoke-GraphWithRetry -Method DELETE `
                        -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($existingPlan.id)" `
                        -Headers @{ "If-Match" = $planToDelete.'@odata.etag' }
                    Write-PlannerLog "  Vorhandener Plan gelöscht." "OK"
                    $script:groupPlansCache.Remove($groupId)  # Cache invalidieren
                    # Neu erstellen
                    $newPlan = Invoke-GraphWithRetry -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/planner/plans" `
                        -Body @{ owner = $groupId; title = $planTitle }
                    Write-PlannerLog "  Plan neu erstellt: $($newPlan.id)" "OK"
                    $script:errorTracker.Plans.Succeeded++
                }
                catch {
                    Add-ErrorToTracker -ItemType "Plan" -ItemName $planTitle -Exception $_ -Context "Überschreiben des Plans"
                    Write-PlannerLog "Kritischer Fehler beim Überschreiben des Plans: $_" "ERROR"
                    return $null
                }
            }
            'v' {
                Write-PlannerLog "  Verwende vorhandenen Plan: $($existingPlan.id)" "OK"
                $newPlan = $existingPlan
                $script:errorTracker.Plans.Succeeded++
            }
            default {
                Write-PlannerLog "  Plan '$planTitle' wird übersprungen." "WARN"
                return $null
            }
        }
    }
    else {
        # Kein vorhandener Plan - normal erstellen
        try {
            $newPlan = Invoke-GraphWithRetry -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/planner/plans" `
                -Body @{ owner = $groupId; title = $planTitle }
            Write-PlannerLog "  Plan erstellt: $($newPlan.id)" "OK"
            $script:errorTracker.Plans.Succeeded++
        }
        catch {
            Add-ErrorToTracker -ItemType "Plan" -ItemName $planTitle -Exception $_ -Context "Erstellen des Plans"
            Write-PlannerLog "Kritischer Fehler: Plan konnte nicht erstellt werden: $_" "ERROR"
            return $null
        }
    }

    # 2. Kategorien/Labels setzen
    if ($planData.Categories) {
        try {
            # Plan-Details abrufen um ETag zu bekommen
            $newPlanDetails = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($newPlan.id)/details"

            $categoryBody = @{
                categoryDescriptions = @{}
            }
            $planData.Categories.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object {
                $categoryBody.categoryDescriptions[$_.Name] = $_.Value
            }

            $params = @{
                Method      = "PATCH"
                Uri         = "https://graph.microsoft.com/v1.0/planner/plans/$($newPlan.id)/details"
                Body        = ($categoryBody | ConvertTo-Json -Depth 10)
                ContentType = "application/json"
                Headers     = @{ "If-Match" = $newPlanDetails.'@odata.etag' }
                OutputType  = "PSObject"
            }
            Invoke-MgGraphRequest @params
            Write-PlannerLog "  Kategorien gesetzt" "OK"
        }
        catch {
            Write-PlannerLog "  Fehler beim Setzen der Kategorien: $_" "WARN"
        }
    }

    # 3. Buckets erstellen (ID-Mapping für Tasks)
    $bucketMapping = @{}
    foreach ($bucket in ($planData.Buckets | Sort-Object orderHint)) {
        $script:errorTracker.Buckets.Attempted++
        try {
            $newBucket = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/buckets" -Body @{
                name    = $bucket.name
                planId  = $newPlan.id
            }
            $bucketMapping[$bucket.id] = $newBucket.id
            Write-PlannerLog "  Bucket erstellt: $($bucket.name)" "OK"
            $script:errorTracker.Buckets.Succeeded++
        }
        catch {
            Add-ErrorToTracker -ItemType "Bucket" -ItemName $bucket.name -Exception $_ -Context "Plan: $planTitle"
            Write-PlannerLog "  Fehler beim Erstellen von Bucket '$($bucket.name)': $_" "ERROR"
        }
    }

    # 4. Tasks erstellen
    $taskMapping = @{}
    $taskCounter = 0
    $totalTasks = $planData.Tasks.Count
    $warningCategoryKey = $null      # wird bei erstem Problem lazy initialisiert
    $warningCategoryResolved = $false

    # UserMap als Hashtable aufbereiten
    $userMap = @{}
    if ($planData.UserMap) {
        $planData.UserMap.PSObject.Properties | ForEach-Object {
            $userMap[$_.Name] = $_.Value
        }
    }

    foreach ($task in $planData.Tasks) {
        $taskCounter++
        $taskHasIssues = $false
        Write-Progress -Activity "Importiere Tasks für '$planTitle'" -Status "Task $taskCounter von ${totalTasks}: $($task.title)" -PercentComplete (($taskCounter / [Math]::Max(1, $totalTasks)) * 100)

        # Abgeschlossene Tasks überspringen wenn gewünscht
        if ($SkipCompletedTasks -and $task.percentComplete -eq 100) {
            Write-PlannerLog "  Task übersprungen (abgeschlossen): $($task.title)"
            continue
        }

        $script:errorTracker.Tasks.Attempted++

        $newBucketId = if ($task.bucketId -and $bucketMapping.ContainsKey($task.bucketId)) {
            $bucketMapping[$task.bucketId]
        } else {
            $null
        }

        # Task-Body aufbauen
        $taskBody = @{
            planId          = $newPlan.id
            title           = $task.title
            percentComplete = $task.percentComplete
            priority        = $task.priority
        }

        if ($newBucketId) {
            $taskBody["bucketId"] = $newBucketId
        }

        $convertedDue   = if ($task.dueDateTime)   { ConvertTo-IsoDate $task.dueDateTime }   else { $null }
        $convertedStart = if ($task.startDateTime) { ConvertTo-IsoDate $task.startDateTime } else { $null }

        # Daten nur setzen wenn gültig: DueDate darf nicht vor StartDate liegen
        if ($convertedStart -and $convertedDue -and ([datetime]$convertedDue -lt [datetime]$convertedStart)) {
            Write-PlannerLog "  Warnung: DueDate ($convertedDue) liegt vor StartDate ($convertedStart) bei Task '$($task.title)' - Daten werden ignoriert" "WARN"
            $taskHasIssues = $true
        } else {
            if ($convertedDue)   { $taskBody["dueDateTime"]   = $convertedDue }
            if ($convertedStart) { $taskBody["startDateTime"] = $convertedStart }
        }

        # Labels/Kategorien
        if ($task.appliedCategories) {
            $categories = @{}
            $task.appliedCategories.PSObject.Properties | Where-Object { $_.Value -eq $true } | ForEach-Object {
                $categories[$_.Name] = $true
            }
            if ($categories.Count -gt 0) {
                $taskBody["appliedCategories"] = $categories
            }
        }

        # Zuweisungen
        if (-not $SkipAssignments -and $task.assignments) {
            $assignments = @{}
            $task.assignments.PSObject.Properties | ForEach-Object {
                $resolvedId = Resolve-UserId -OldUserId $_.Name -OldUserMap $userMap
                if ($resolvedId) {
                    $assignments[$resolvedId] = @{
                        "@odata.type"  = "#microsoft.graph.plannerAssignment"
                        "orderHint"    = " !"
                    }
                }
                else {
                    $userName = if ($userMap[$_.Name]) { $userMap[$_.Name].DisplayName } else { $_.Name }
                    $userUPN = if ($userMap[$_.Name]) { $userMap[$_.Name].UserPrincipalName } else { 'Unbekannt' }
                    Write-PlannerLog "    Benutzer konnte nicht zugewiesen werden: $userName" "WARN"
                    $taskHasIssues = $true
                    
                    # Für späteres Reporting sammeln
                    $missingUsers.Add([PSCustomObject]@{
                        PlanName = $planTitle
                        BucketName = $bucket.name
                        TaskTitle = $task.title
                        TaskId = $newTask.id
                        FehlenderBenutzer = $userName
                        BenutzerUPN = $userUPN
                        BenutzerAltId = $_.Name
                    })
                }
            }
            if ($assignments.Count -gt 0) {
                $taskBody["assignments"] = $assignments
            }
        }

        try {
            $newTask = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/tasks" -Body $taskBody
            $taskMapping[$task.id] = $newTask.id
            Write-PlannerLog "  Task erstellt: $($task.title)" "OK"
            $script:errorTracker.Tasks.Succeeded++

            # 5. Task-Details setzen (Beschreibung, Checkliste, Referenzen)
            $detail = $planData.TaskDetails | Where-Object { $_.taskId -eq $task.id }
            if ($detail) {
                $script:errorTracker.TaskDetails.Attempted++
                $hasDetails = $false
                $detailBody = @{}

                # Beschreibung
                if ($detail.description) {
                    $detailBody["description"] = $detail.description
                    $detailBody["previewType"] = "description"
                    $hasDetails = $true
                }

                # Checkliste
                if ($detail.checklist) {
                    $checklist = @{}
                    $detail.checklist.PSObject.Properties | ForEach-Object {
                        $checkId = [Guid]::NewGuid().ToString()
                        $checklist[$checkId] = @{
                            "@odata.type" = "microsoft.graph.plannerChecklistItem"
                            title         = $_.Value.title
                            isChecked     = $_.Value.isChecked
                        }
                    }
                    if ($checklist.Count -gt 0) {
                        $detailBody["checklist"] = $checklist
                        $hasDetails = $true
                    }
                }

                # Referenzen/Links
                if ($detail.references) {
                    $references = @{}
                    $detail.references.PSObject.Properties | ForEach-Object {
                        $url = $_.Name
                        $references[$url] = @{
                            "@odata.type" = "microsoft.graph.plannerExternalReference"
                            alias         = $_.Value.alias
                            type          = $_.Value.type
                        }
                        # previewPriority absichtlich nicht übernehmen: exportierte Werte
                        # sind mandantenspezifische orderHints und im Ziel-Tenant ungültig
                    }
                    if ($references.Count -gt 0) {
                        $detailBody["references"] = $references
                        $hasDetails = $true
                    }
                }

                if ($hasDetails) {
                    try {
                        # ETag für Task-Details holen
                        $newTaskDetails = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)/details"

                        $patchParams = @{
                            Method      = "PATCH"
                            Uri         = "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)/details"
                            Body        = ($detailBody | ConvertTo-Json -Depth 20)
                            ContentType = "application/json"
                            Headers     = @{ "If-Match" = $newTaskDetails.'@odata.etag' }
                            OutputType  = "PSObject"
                        }
                        Invoke-MgGraphRequest @patchParams
                        Write-PlannerLog "    Details gesetzt für: $($task.title)" "OK"
                        $script:errorTracker.TaskDetails.Succeeded++
                    }
                    catch {
                        Add-ErrorToTracker -ItemType "TaskDetail" -ItemName $task.title -Exception $_ -Context "Details für Task"
                        Write-PlannerLog "    Fehler beim Setzen der Task-Details: $_" "WARN"
                        $taskHasIssues = $true
                    }
                }
            }

            # Warn-Label setzen wenn es Probleme gab
            if ($taskHasIssues) {
                # Warn-Kategorie lazy initialisieren (nur beim ersten Problem)
                if (-not $warningCategoryResolved) {
                    $warningCategoryResolved = $true
                    $warningCategoryKey = Get-OrCreateWarningCategory -PlanId $newPlan.id -ExistingCategories $planData.Categories
                }
                if ($warningCategoryKey) {
                    try {
                        $taskForEtag = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)"
                        
                        # Hashtable explizit aufbauen (Variable als Key)
                        $bodyObj = @{ appliedCategories = @{} }
                        $bodyObj.appliedCategories[$warningCategoryKey] = $true
                        
                        # JSON-Body erstellen
                        $jsonBody = $bodyObj | ConvertTo-Json -Depth 5 -Compress
                        
                        $warnParams = @{
                            Method      = "PATCH"
                            Uri         = "https://graph.microsoft.com/v1.0/planner/tasks/$($newTask.id)"
                            Body        = $jsonBody
                            ContentType = "application/json"
                            Headers     = @{ "If-Match" = $taskForEtag.'@odata.etag' }
                            OutputType  = "PSObject"
                        }
                        Invoke-MgGraphRequest @warnParams
                        Write-PlannerLog "    Warn-Label gesetzt für: $($task.title)" "WARN"
                    }
                    catch {
                        Write-PlannerLog "    [DEBUG] Warn-Label Fehler Details:" "WARN"
                        Write-PlannerLog "      - CategoryKey: '$warningCategoryKey' (Typ: $($warningCategoryKey.GetType().Name))" "WARN"
                        Write-PlannerLog "      - BodyObj.appliedCategories Keys: $($bodyObj.appliedCategories.Keys -join ', ')" "WARN"
                        Write-PlannerLog "      - BodyObj Typ: $($bodyObj.GetType().Name)" "WARN"
                        Write-PlannerLog "      - BodyObj.appliedCategories Typ: $($bodyObj.appliedCategories.GetType().Name)" "WARN"
                        Write-PlannerLog "      - JSON Body: $jsonBody" "WARN"
                        Write-PlannerLog "      - Fehlermeldung: $_" "WARN"
                        Write-PlannerLog "      - Exception Typ: $($_.Exception.GetType().FullName)" "WARN"
                        if ($_.Exception.InnerException) {
                            Write-PlannerLog "      - Inner Exception: $($_.Exception.InnerException.Message)" "WARN"
                        }
                    }
                }
            }
        }
        catch {
            Add-ErrorToTracker -ItemType "Task" -ItemName $task.title -Exception $_ -Context "Plan: $planTitle"
            Write-PlannerLog "  Fehler beim Erstellen von Task '$($task.title)': $_" "ERROR"
        }
    }

    Write-Progress -Activity "Importiere Tasks" -Completed

    # Import-Mapping speichern (für Referenz)
    $mappingData = @{
        ImportDate   = (Get-Date).ToString("o")
        OriginalPlan = $planData.Plan.id
        NewPlanId    = $newPlan.id
        GroupId      = $groupId
        BucketMap    = $bucketMapping
        TaskMap      = $taskMapping
    }
    $planFileName = ($planTitle -replace '[\\/:*?"<>|]', '_')
    try {
        $mappingData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$ImportPath\${planFileName}_ImportMapping.json" -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-PlannerLog "Fehler beim Schreiben der Mapping-Datei: $_" "ERROR"
    }
    
    # Fehlende Benutzer in separate Datei schreiben
    if ($missingUsers.Count -gt 0) {
        $missingUsersFile = "$ImportPath\${planFileName}_Fehlende_Benutzer.csv"
        try {
            # CSV mit direkten Task-Links erstellen
            $csvContent = [System.Collections.Generic.List[string]]::new()
            $csvContent.Add('"Plan";"Bucket";"Task";"Fehlender Benutzer";"Benutzer UPN";"Task-Link"')
            
            foreach ($item in $missingUsers) {
                $taskUrl = "https://tasks.office.com/" + $groupId + "/Home/Task/" + $item.TaskId
                $line = '"{0}";"{1}";"{2}";"{3}";"{4}";"{5}"' -f `
                    $item.PlanName, `
                    $item.BucketName, `
                    $item.TaskTitle, `
                    $item.FehlenderBenutzer, `
                    $item.BenutzerUPN, `
                    $taskUrl
                $csvContent.Add($line)
            }
            
            $csvContent | Out-File -FilePath $missingUsersFile -Encoding UTF8 -Force
            Write-PlannerLog "Fehlende Benutzer protokolliert: $missingUsersFile ($($missingUsers.Count) Einträge)" "WARN"
        }
        catch {
            Write-PlannerLog "Fehler beim Schreiben der fehlenden Benutzer: $_" "ERROR"
        }
    }

    return @{
        NewPlanId      = $newPlan.id
        PlanTitle      = $planTitle
        GroupName      = $planData.Plan.groupDisplayName
        GroupId        = $groupId
        TasksCreated   = $taskMapping.Count
        BucketsCreated = $bucketMapping.Count
    }
}

function Get-M365GroupsForListing {
    <#
    .SYNOPSIS
        Ruft alle M365-Gruppen (Unified) aus dem Tenant ab und gibt sie sortiert aus.
    #>
    $groups = [System.Collections.Generic.List[PSCustomObject]]::new()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(g:g eq 'Unified')" +
           "&`$select=id,displayName,mail&`$top=999"
    try {
        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
            if ($response.value) {
                foreach ($g in $response.value) {
                    $groups.Add([PSCustomObject]@{
                        Id          = $g.id
                        DisplayName = $g.displayName
                        Mail        = $g.mail
                    })
                }
            }
            $uri = $response.'@odata.nextLink'
        } while ($uri)
    }
    catch {
        Write-PlannerLog "Fehler beim Abrufen der Gruppen: $_" "ERROR"
        return $null
    }
    return @($groups | Sort-Object DisplayName)
}

function Find-PlannerExports {
    <#
    .SYNOPSIS
        Durchsucht Standard-Speicherorte nach Planner-Export-Verzeichnissen.
    .OUTPUTS
        Array von PSCustomObject mit Export-Metadaten, sortiert nach Datum (neueste zuerst).
    #>

    # Standard-Suchpfade (gleiche Priorität wie beim Export)
    $searchPaths = [System.Collections.Generic.List[string]]::new()

    # Aktuelles Verzeichnis
    $searchPaths.Add((Get-Location).Path)

    # Standard-Exportpfade
    if (Test-Path 'C:\temp' -PathType Container) { $searchPaths.Add('C:\temp') }
    if (Test-Path 'C:\tmp'  -PathType Container) { $searchPaths.Add('C:\tmp') }
    $docsPath = [System.Environment]::GetFolderPath('MyDocuments')
    if ($docsPath -and (Test-Path $docsPath -PathType Container)) { $searchPaths.Add($docsPath) }

    $foundExports = [System.Collections.Generic.List[PSCustomObject]]::new()
    $seenPaths    = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    foreach ($searchPath in $searchPaths) {
        $exportDirs = Get-ChildItem -Path $searchPath -Directory -Filter 'PlannerExport_*' `
            -Recurse -Depth 1 -ErrorAction SilentlyContinue
        foreach ($dir in $exportDirs) {
            if (-not $seenPaths.Add($dir.FullName)) { continue }

            $indexFile = Join-Path $dir.FullName '_ExportIndex.json'
            if (-not (Test-Path $indexFile)) { continue }

            try {
                $index = Get-Content $indexFile -Raw -Encoding UTF8 | ConvertFrom-Json
                $foundExports.Add([PSCustomObject]@{
                    Path       = $dir.FullName
                    FolderName = $dir.Name
                    ExportDate = $index.ExportDate
                    ExportedBy = $index.ExportedBy
                    TotalPlans = $index.TotalPlans
                    Plans      = $index.Plans
                })
            }
            catch {
                # Index-Datei nicht lesbar – trotzdem aufnehmen
                $foundExports.Add([PSCustomObject]@{
                    Path       = $dir.FullName
                    FolderName = $dir.Name
                    ExportDate = $null
                    ExportedBy = $null
                    TotalPlans = 0
                    Plans      = @()
                })
            }
        }
    }

    # Neueste Exporte zuerst
    return @($foundExports | Sort-Object {
        if ($_.ExportDate) {
            try { [datetime]$_.ExportDate } catch { [datetime]::MinValue }
        } else { [datetime]::MinValue }
    } -Descending)
}

function Show-ExportSelectionMenu {
    <#
    .SYNOPSIS
        Zeigt ein interaktives Auswahlmenü für gefundene Planner-Exporte.
    .PARAMETER Exports
        Array von Export-Objekten (Ausgabe von Find-PlannerExports).
    .OUTPUTS
        Pfad des gewählten Exports oder $null bei Abbruch.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$Exports
    )

    Write-Host ''
    Write-Host 'Verfügbare Planner-Exporte:' -ForegroundColor Yellow
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray

    for ($i = 0; $i -lt $Exports.Count; $i++) {
        $exp     = $Exports[$i]
        $dateStr = if ($exp.ExportDate) {
            try { ([datetime]$exp.ExportDate).ToString('dd.MM.yyyy HH:mm:ss') } catch { $exp.ExportDate }
        } else { 'Datum unbekannt' }

        Write-Host ''
        Write-Host "  [$($i + 1)] $($exp.FolderName)" -ForegroundColor Cyan
        Write-Host "      Exportiert am  : $dateStr" -ForegroundColor Gray
        if ($exp.ExportedBy) {
            Write-Host "      Exportiert von : $($exp.ExportedBy)" -ForegroundColor Gray
        }

        if ($exp.Plans -and $exp.Plans.Count -gt 0) {
            foreach ($plan in $exp.Plans) {
                $groupInfo = if ($plan.GroupName) { " (Gruppe: $($plan.GroupName))" } else { '' }
                $taskInfo  = if ($null -ne $plan.Tasks) { ", $($plan.Tasks) Tasks" } else { '' }
                $bucketInfo = if ($null -ne $plan.Buckets) { ", $($plan.Buckets) Buckets" } else { '' }
                Write-Host "        - $($plan.PlanTitle)$groupInfo$taskInfo$bucketInfo" -ForegroundColor White
            }
        } else {
            Write-Host '        (Keine Plan-Metadaten verfügbar)' -ForegroundColor DarkGray
        }
    }

    Write-Host ''
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    do {
        $selection = Read-Host "Export auswählen [1-$($Exports.Count)] oder 'q' zum Abbrechen"
        if ($selection -eq 'q' -or $selection -eq 'Q') {
            return $null
        }
        $num = 0
        if ([int]::TryParse($selection, [ref]$num) -and $num -ge 1 -and $num -le $Exports.Count) {
            return $Exports[$num - 1].Path
        }
        Write-Host "  Ungültige Eingabe. Bitte eine Zahl zwischen 1 und $($Exports.Count) eingeben." -ForegroundColor Red
    } while ($true)
}

#endregion

#region Hauptprogramm

# Pro Programmstart eine eigene Log-Datei (ohne Farbcodes)
$logsDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir | Out-Null }
$transcriptPath = Join-Path $logsDir ("import_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText
}
Start-Transcript -Path $transcriptPath -NoClobber -ErrorAction SilentlyContinue

try {

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Planner Import Tool" -ForegroundColor Cyan
Write-Host "  by Alexander Waller" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "  *** DRY RUN MODUS - Es werden keine Änderungen vorgenommen ***" -ForegroundColor Magenta
    Write-Host ""
}

# Wenn kein ImportPath angegeben: Standard-Speicherorte nach Exporten durchsuchen
# (kein Graph-Login nötig – liest nur lokale JSON-Dateien)
# Bei -ListGroups wird dieser Block übersprungen
if ([string]::IsNullOrEmpty($ImportPath) -and -not $ListGroups) {
    Write-Host 'Suche nach Planner-Exporten in Standard-Speicherorten...' -ForegroundColor Yellow
    Write-Host ''

    $availableExports = Find-PlannerExports

    if ($availableExports.Count -eq 0) {
        Write-Host 'Keine Planner-Exporte gefunden.' -ForegroundColor Red
        Write-Host ''
        Write-Host 'Durchsuchte Verzeichnisse:' -ForegroundColor White
        Write-Host '  - Aktuelles Verzeichnis' -ForegroundColor Gray
        if (Test-Path 'C:\temp' -PathType Container) { Write-Host '  - C:\temp' -ForegroundColor Gray }
        if (Test-Path 'C:\tmp'  -PathType Container) { Write-Host '  - C:\tmp'  -ForegroundColor Gray }
        $docsPath = [System.Environment]::GetFolderPath('MyDocuments')
        if ($docsPath) { Write-Host "  - $docsPath" -ForegroundColor Gray }
        Write-Host ''
        Write-Host 'Import-Pfad manuell angeben:' -ForegroundColor White
        Write-Host '  .\Import-PlannerData.ps1 -ImportPath "<Pfad-zum-Export-Verzeichnis>"' -ForegroundColor White
        Write-Host ''
        exit 1
    }

    # DryRun ohne ImportPath: alle Exporte mit Gruppen auflisten, kein Import
    if ($DryRun) {
        Write-Host '============================================================' -ForegroundColor Cyan
        Write-Host '  ÜBERSICHT ALLER VERFÜGBAREN EXPORTE (DRY RUN)' -ForegroundColor Cyan
        Write-Host '============================================================' -ForegroundColor Cyan

        $totalPlans   = 0
        $totalTasks   = 0
        $totalBuckets = 0

        foreach ($exp in $availableExports) {
            $dateStr = if ($exp.ExportDate) {
                try { ([datetime]$exp.ExportDate).ToString('dd.MM.yyyy HH:mm:ss') } catch { $exp.ExportDate }
            } else { 'Datum unbekannt' }

            Write-Host ''
            Write-Host "  $($exp.FolderName)" -ForegroundColor Cyan
            Write-Host "    Speicherort: $($exp.Path)" -ForegroundColor Gray
            Write-Host "    Exportiert am: $dateStr" -ForegroundColor Gray
            if ($exp.ExportedBy) {
                Write-Host "    Exportiert von: $($exp.ExportedBy)" -ForegroundColor Gray
            }

            if ($exp.Plans -and $exp.Plans.Count -gt 0) {
                # Gruppen zusammenfassen
                $groups = @($exp.Plans | Where-Object { $_.GroupName } |
                    Select-Object -ExpandProperty GroupName -Unique | Sort-Object)
                if ($groups.Count -gt 0) {
                    Write-Host "    Gruppen: $($groups -join ', ')" -ForegroundColor White
                }

                Write-Host "    Pläne ($($exp.Plans.Count)):" -ForegroundColor White
                foreach ($plan in $exp.Plans) {
                    $grp     = if ($plan.GroupName) { " (Gruppe: $($plan.GroupName))" } else { '' }
                    $tasks   = if ($null -ne $plan.Tasks)   { "$($plan.Tasks) Tasks"   } else { '? Tasks' }
                    $buckets = if ($null -ne $plan.Buckets) { "$($plan.Buckets) Buckets" } else { '? Buckets' }
                    Write-Host "      - $($plan.PlanTitle)$grp  |  $buckets, $tasks" -ForegroundColor Gray
                    $totalPlans++
                    $totalTasks   += [int]($plan.Tasks)
                    $totalBuckets += [int]($plan.Buckets)
                }
            } else {
                Write-Host '    (Keine Plan-Metadaten verfügbar)' -ForegroundColor DarkGray
            }
        }

        Write-Host ''
        Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
        Write-Host "  Gesamt: $($availableExports.Count) Export(e)  |  $totalPlans Pläne  |  $totalBuckets Buckets  |  $totalTasks Tasks" -ForegroundColor Yellow
        Write-Host ''
        Write-Host 'Import starten mit:' -ForegroundColor White
        Write-Host '  .\Import-PlannerData.ps1 -ImportPath "<Pfad>"' -ForegroundColor Gray
        Write-Host '  .\Import-PlannerData.ps1   (interaktive Auswahl)' -ForegroundColor Gray
        Write-Host ''
        exit 0
    }

    $selectedPath = Show-ExportSelectionMenu -Exports $availableExports

    if ($null -eq $selectedPath) {
        Write-Host ''
        Write-Host 'Import abgebrochen.' -ForegroundColor Yellow
        exit 0
    }

    $ImportPath = $selectedPath
    Write-Host ''
    Write-Host "Gewähltes Export-Verzeichnis: $ImportPath" -ForegroundColor Green
    Write-Host ''
}

# ListGroups-Modus: Gruppen auflisten und beenden (benötigt keinen ImportPath)
if ($ListGroups) {
    Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    
    if (-not (Connect-ToGraph)) {
        Write-PlannerLog "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
        exit 1
    }
    
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  M365-GRUPPEN IM TENANT' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''

    $groups = Get-M365GroupsForListing
    if ($null -eq $groups) {
        exit 1
    }
    if ($groups.Count -eq 0) {
        Write-Host '  Keine M365-Gruppen gefunden.' -ForegroundColor Yellow
        Write-Host ''
        exit 0
    }

    $maxLen = ($groups | ForEach-Object { $_.DisplayName.Length } | Measure-Object -Maximum).Maximum
    $maxLen = [Math]::Max($maxLen, 20)

    Write-Host (('  {0,-' + $maxLen + '}  {1,-38}  {2}') -f 'Gruppenname', 'Gruppen-ID', 'E-Mail') -ForegroundColor Yellow
    Write-Host (('  {0}  {1}  {2}') -f ('-' * $maxLen), ('-' * 38), ('-' * 40)) -ForegroundColor DarkGray

    foreach ($g in $groups) {
        Write-Host (('  {0,-' + $maxLen + '}  {1,-38}  {2}') -f $g.DisplayName, $g.Id, $g.Mail)
    }

    Write-Host ''
    Write-Host "  Gesamt: $($groups.Count) Gruppe(n)" -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Tipp: Gruppen-ID als Ziel angeben mit:' -ForegroundColor White
    Write-Host '  .\Import-PlannerData.ps1 -ImportPath "<Pfad>" -TargetGroupId "<ID>"' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Graph-Verbindung bleibt aktiv für nachfolgende Befehle.' -ForegroundColor DarkGray
    Write-Host ''

    exit 0
}

# Validiere Import-Verzeichnis (nur wenn nicht ListGroups)
$pathValidationError = $null
if (-not (Test-SafePath -Path $ImportPath -Mode Import -ErrorMessage ([ref]$pathValidationError))) {
    Write-Host ""
    Write-Host "Fehler: $pathValidationError" -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Lade Index-Datei
$indexPath = Join-Path $ImportPath "_ExportIndex.json"
$indexData = $null
if (Test-Path $indexPath) {
    $indexData = Get-Content $indexPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

# JSON-Dateien finden
$jsonFiles = Get-ChildItem -Path $ImportPath -Filter "*.json" |
    Where-Object { $_.Name -ne "_ExportIndex.json" -and $_.Name -notmatch "ImportMapping" -and $_.Name -ne "import_errors.json" }

if ($jsonFiles.Count -eq 0) {
    Write-Host "" -ForegroundColor Red
    Write-Host "Keine Export-Dateien gefunden in: $ImportPath" -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Im DryRun-Modus: Zeige detaillierte Übersicht der zu importierenden Daten
if ($DryRun) {
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host '  ÜBERSICHT DER ZU IMPORTIERENDEN DATEN (DRY RUN)' -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host "Speicherort: $ImportPath" -ForegroundColor Gray
    
    if ($indexData) {
        $dateStr = if ($indexData.ExportDate) {
            try { ([datetime]$indexData.ExportDate).ToString('dd.MM.yyyy HH:mm:ss') } catch { $indexData.ExportDate }
        } else { 'Datum unbekannt' }
        Write-Host "Exportiert am: $dateStr" -ForegroundColor Gray
        if ($indexData.ExportedBy) {
            Write-Host "Exportiert von: $($indexData.ExportedBy)" -ForegroundColor Gray
        }
    }
    
    Write-Host ''
    Write-Host "Gefundene Pläne ($($jsonFiles.Count)):" -ForegroundColor White
    Write-Host ''
    
    $totalTasks = 0
    $totalBuckets = 0
    $totalCompletedTasks = 0
    $totalTasksWithAssignments = 0
    
    foreach ($jsonFile in $jsonFiles) {
        try {
            $planData = Get-Content $jsonFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $planTitle = if ($planData.Plan.title) { $planData.Plan.title } else { $jsonFile.BaseName }
            $groupName = if ($planData.Plan.groupDisplayName) { $planData.Plan.groupDisplayName } else { 'Gruppe unbekannt' }
            $groupId = if ($planData.Plan.groupId) { $planData.Plan.groupId } else { 'ID unbekannt' }
            
            $taskCount = if ($planData.Tasks) { $planData.Tasks.Count } else { 0 }
            $bucketCount = if ($planData.Buckets) { $planData.Buckets.Count } else { 0 }
            $completedTasks = if ($planData.Tasks) { @($planData.Tasks | Where-Object { $_.percentComplete -eq 100 }).Count } else { 0 }
            $tasksWithAssignments = if ($planData.Tasks) { @($planData.Tasks | Where-Object { $_.assignments -and $_.assignments.PSObject.Properties.Count -gt 0 }).Count } else { 0 }
            
            Write-Host "  - $planTitle" -ForegroundColor Cyan
            Write-Host "      Gruppe: $groupName" -ForegroundColor Gray
            Write-Host "      Gruppen-ID: $groupId" -ForegroundColor DarkGray
            Write-Host "      Buckets: $bucketCount" -ForegroundColor White
            Write-Host "      Tasks: $taskCount" -ForegroundColor White
            
            if ($completedTasks -gt 0) {
                Write-Host "        davon abgeschlossen: $completedTasks" -ForegroundColor Gray
            }
            if ($tasksWithAssignments -gt 0) {
                Write-Host "        davon mit Zuweisungen: $tasksWithAssignments" -ForegroundColor Gray
            }
            
            $totalTasks += $taskCount
            $totalBuckets += $bucketCount
            $totalCompletedTasks += $completedTasks
            $totalTasksWithAssignments += $tasksWithAssignments
            
            Write-Host ''
        }
        catch {
            Write-Host "  - $($jsonFile.Name)" -ForegroundColor Yellow
            Write-Host "      (Fehler beim Lesen der Datei)" -ForegroundColor DarkGray
            Write-Host ''
        }
    }
    
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host "Gesamt: $($jsonFiles.Count) Plan/Pläne  |  $totalBuckets Buckets  |  $totalTasks Tasks" -ForegroundColor Yellow
    if ($totalCompletedTasks -gt 0) {
        Write-Host "  davon abgeschlossen: $totalCompletedTasks Tasks" -ForegroundColor Gray
    }
    if ($totalTasksWithAssignments -gt 0) {
        Write-Host "  davon mit Zuweisungen: $totalTasksWithAssignments Tasks" -ForegroundColor Gray
    }
    Write-Host ''
    
    if ($TargetGroupId) {
        Write-Host "Zielgruppe: $TargetGroupId" -ForegroundColor White
    } else {
        Write-Host "Zielgruppe: Originalgruppe aus Export" -ForegroundColor White
    }
    
    if ($SkipAssignments) {
        Write-Host "Option: Zuweisungen werden übersprungen" -ForegroundColor Yellow
    }
    if ($SkipCompletedTasks) {
        Write-Host "Option: Abgeschlossene Tasks werden übersprungen" -ForegroundColor Yellow
    }
    if ($UserMapping -and $UserMapping.Count -gt 0) {
        Write-Host "Option: Benutzer-Mapping aktiv ($($UserMapping.Count) Zuordnung(en))" -ForegroundColor Yellow
    }
    
    Write-Host ''
    Write-Host '------------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host 'Import starten mit:' -ForegroundColor White
    Write-Host "  .\Import-PlannerData.ps1 -ImportPath \"$ImportPath\"" -ForegroundColor Gray
    Write-Host ''
    
    exit 0
}

# Microsoft Graph Module und Verbindung (nur wenn nicht DryRun)
Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

if (-not (Connect-ToGraph)) {
    Write-PlannerLog "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
    exit 1
}

# JSON-Dateien wurden bereits oben geladen und validiert
Write-PlannerLog "Import-Verzeichnis: $ImportPath"

if ($indexData) {
    Write-PlannerLog "Export-Index geladen. Export vom: $($indexData.ExportDate)"
    Write-PlannerLog "Exportiert von: $($indexData.ExportedBy)"
    Write-PlannerLog "Pläne im Export: $($indexData.TotalPlans)"
}

Write-Host ""
Write-Host "Gefundene Export-Dateien:" -ForegroundColor Yellow
for ($i = 0; $i -lt $jsonFiles.Count; $i++) {
    Write-Host "  [$($i+1)] $($jsonFiles[$i].Name)"
}
Write-Host ""

# Bestätigung
if (-not $DryRun) {
    $confirm = Read-Host "Möchten Sie alle $($jsonFiles.Count) Pläne importieren? (j/n)"
    if ($confirm -ne 'j' -and $confirm -ne 'J') {
        Write-PlannerLog "Import abgebrochen durch Benutzer." "WARN"
        exit 0
    }
}

# Validiere Zielgruppe wenn angegeben
if ($TargetGroupId) {
    Write-Host ""
    if (-not (Test-TargetGroup -GroupId $TargetGroupId)) {
        Write-PlannerLog "Abbruch: Zielgruppe ist ungültig oder nicht zugreifbar." "ERROR"
        exit 1
    }
    Write-Host ""
}

# Import durchführen
$importResults = @()
foreach ($jsonFile in $jsonFiles) {
    Write-Host ""
    Write-PlannerLog "=== Importiere: $($jsonFile.Name) ===" "OK"

    $result = Import-PlanFromJson -JsonFilePath $jsonFile.FullName -TargetGroupId $TargetGroupId
    if ($result) {
        $importResults += $result
    }
}

# Zusammenfassung
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  IMPORT ABGESCHLOSSEN" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

if ($DryRun) {
    Write-Host "  *** Dies war ein DRY RUN - keine Änderungen wurden vorgenommen ***" -ForegroundColor Magenta
    Write-Host ""
    
    # DryRun summary
    $validResults = @($importResults | Where-Object { $_.Status -eq "DryRun" })
    if ($validResults.Count -gt 0) {
        $totalTasksToImport   = ($validResults | ForEach-Object { [int]($_.TasksCreated)   } | Measure-Object -Sum).Sum
        $totalBucketsToImport = ($validResults | ForEach-Object { [int]($_.BucketsCreated) } | Measure-Object -Sum).Sum
        Write-Host "  Würde importieren:" -ForegroundColor White
        Write-Host "    Pläne:    $($validResults.Count)" -ForegroundColor Cyan
        Write-Host "    Buckets:  $totalBucketsToImport" -ForegroundColor Cyan
        Write-Host "    Tasks:    $totalTasksToImport" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Details je Plan:" -ForegroundColor White
        foreach ($r in $validResults) {
            $grp = if ($r.GroupName) { " (Gruppe: $($r.GroupName))" } else { " (Gruppe-ID: $($r.TargetGroupId))" }
            Write-Host "    - $($r.PlanTitle)$grp" -ForegroundColor Cyan
            Write-Host "        Buckets: $($r.BucketsCreated)   Tasks: $($r.TasksCreated)" -ForegroundColor Gray
        }
        Write-Host ""
    }

    $failedValidations = $importResults | Where-Object { $null -eq $_.Status -or $_.Status -ne "DryRun" }
    if ($failedValidations.Count -gt 0) {
        Write-Host "  Validierungsfehler bei $($failedValidations.Count) Plan(en)" -ForegroundColor Yellow
        Write-Host ""
    }
}
else {
    # Success summary
    $totalTasksImported   = ($importResults | ForEach-Object { [int]($_.TasksCreated)   } | Measure-Object -Sum).Sum
    $totalBucketsImported = ($importResults | ForEach-Object { [int]($_.BucketsCreated) } | Measure-Object -Sum).Sum
    Write-Host "  Erfolgreich importiert:" -ForegroundColor White
    Write-Host "    Pläne:    $($importResults.Count)" -ForegroundColor Green
    Write-Host "    Buckets:  $totalBucketsImported" -ForegroundColor Green
    Write-Host "    Tasks:    $totalTasksImported" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Details je Plan:" -ForegroundColor White
    foreach ($r in $importResults) {
        $grp = if ($r.GroupName) { " (Gruppe: $($r.GroupName))" } else { " (Gruppe-ID: $($r.GroupId))" }
        Write-Host "    - $($r.PlanTitle)$grp" -ForegroundColor Green
        Write-Host "        Buckets: $($r.BucketsCreated)   Tasks: $($r.TasksCreated)" -ForegroundColor Gray
    }
    Write-Host ""

    # Fehlende Benutzer-Zusammenfassung
    $missingUserFiles = Get-ChildItem -Path $ImportPath -Filter "*_Fehlende_Benutzer.csv" -ErrorAction SilentlyContinue
    if ($missingUserFiles -and $missingUserFiles.Count -gt 0) {
        Write-Host ""
        Write-Host "  ⚠ FEHLENDE BENUTZERZUWEISUNGEN" -ForegroundColor Yellow
        Write-Host "  ══════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Bei einigen Tasks konnten Benutzer nicht zugewiesen werden." -ForegroundColor White
        Write-Host "  Details in folgenden Dateien (mit Excel öffnen):" -ForegroundColor White
        Write-Host ""
        
        foreach ($file in $missingUserFiles) {
            $lineCount = (Get-Content $file.FullName -Encoding UTF8 | Measure-Object -Line).Lines - 1  # -1 für Header
            Write-Host "    📄 $($file.Name)" -ForegroundColor Cyan
            Write-Host "       $lineCount fehlende Zuweisung(en)" -ForegroundColor Gray
            Write-Host "       $($file.FullName)" -ForegroundColor DarkGray
            Write-Host ""
        }
        
        Write-Host "  Die CSV-Dateien enthalten:" -ForegroundColor White
        Write-Host "    • Plan- und Task-Namen zur Orientierung" -ForegroundColor Gray
        Write-Host "    • Direkte Links zu den Tasks (klickbar in Excel)" -ForegroundColor Gray
        Write-Host "    • Namen der fehlenden Benutzer" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  Anleitung:" -ForegroundColor White
        Write-Host "    1. CSV-Datei mit Excel öffnen (Trennzeichen: Semikolon)" -ForegroundColor Gray
        Write-Host "    2. Task-Link anklicken → Task öffnet sich in Planner" -ForegroundColor Gray
        Write-Host "    3. Richtigen Benutzer manuell zuweisen" -ForegroundColor Gray
        Write-Host ""
    }

    # Cache statistics
    Write-CacheStatistics
}

# Error summary and exit code
# Im DryRun keine Fehlerstatistik ausgeben (Zähler wurden nicht befüllt)
if ($DryRun) {
    $exitCode = 0
} else {
    $exitCode = Write-ErrorSummary -OutputPath $ImportPath
}

Write-PlannerLog "Import abgeschlossen mit Exit-Code: $exitCode"
# Graph-Verbindung bleibt aktiv für weitere Befehle
# Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

exit $exitCode

} finally {
    Stop-Transcript -ErrorAction SilentlyContinue
}

#endregion
