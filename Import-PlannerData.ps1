<#
.SYNOPSIS
    Importiert zuvor exportierte Microsoft Planner-Daten zurück in neue Pläne.

.DESCRIPTION
    Liest die JSON-Exportdateien von Export-PlannerData.ps1 und erstellt neue Pläne,
    Buckets, Tasks (inkl. Checklisten, Beschreibungen, Zuweisungen und Labels)
    in den angegebenen Microsoft 365 Gruppen.

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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
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
    try {
        $logEntry | Out-File -FilePath "$ImportPath\import.log" -Append -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-Host "[ERROR] Konnte nicht in Log-Datei schreiben: $_" -ForegroundColor Red
    }
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
        Message = if ($Exception) { $Exception.Message } else { "Unbekannter Fehler" }
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
    $itemTypePlural = "$ItemType" + "s"
    $script:errorTracker.$itemTypePlural.Failed += $errorDetails

    # Categorize error
    $errorMessage = $errorDetails.Message.ToLower()
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

function Invoke-GraphWithRetry {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
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
    Write-PlannerLog "Verbinde mit Microsoft Graph..."
    try {
        $context = Get-MgContext
        if ($null -eq $context) {
            try {
                # Versuche zuerst interaktive Anmeldung
                Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -NoWelcome -ErrorAction Stop
            }
            catch {
                # Fallback auf Device Code Flow wenn Browser-Auth fehlschlägt
                Write-PlannerLog "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." "WARN"
                Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -UseDeviceCode -NoWelcome
            }
        }
        $context = Get-MgContext
        if ($null -eq $context -or [string]::IsNullOrEmpty($context.Account)) {
            throw "Keine gültige Verbindung hergestellt"
        }
        Write-PlannerLog "Verbunden als: $($context.Account)" "OK"
        return $true
    }
    catch {
        Write-PlannerLog "Fehler bei der Verbindung: $_" "ERROR"
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
        Write-PlannerLog "[DRY RUN] Würde Plan '$planTitle' erstellen" "DRYRUN"
        Write-PlannerLog "[DRY RUN] Buckets: $($planData.Buckets.Count)" "DRYRUN"
        Write-PlannerLog "[DRY RUN] Tasks: $($planData.Tasks.Count)" "DRYRUN"
        return
    }

    # 1. Plan erstellen
    $script:errorTracker.Plans.Attempted++

    try {
        $newPlan = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/plans" -Body @{
            owner = $groupId
            title = $planTitle
        }
        Write-PlannerLog "  Plan erstellt: $($newPlan.id)" "OK"
        $script:errorTracker.Plans.Succeeded++
    }
    catch {
        Add-ErrorToTracker -ItemType "Plan" -ItemName $planTitle -Exception $_ -Context "Erstellen des Plans"
        Write-PlannerLog "Kritischer Fehler: Plan konnte nicht erstellt werden: $_" "ERROR"
        return $null
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

    # UserMap als Hashtable aufbereiten
    $userMap = @{}
    if ($planData.UserMap) {
        $planData.UserMap.PSObject.Properties | ForEach-Object {
            $userMap[$_.Name] = $_.Value
        }
    }

    foreach ($task in $planData.Tasks) {
        $taskCounter++
        Write-Progress -Activity "Importiere Tasks für '$planTitle'" -Status "Task $taskCounter von ${totalTasks}: $($task.title)" -PercentComplete (($taskCounter / $totalTasks) * 100)

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

        if ($task.dueDateTime) {
            $taskBody["dueDateTime"] = $task.dueDateTime
        }

        if ($task.startDateTime) {
            $taskBody["startDateTime"] = $task.startDateTime
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
                    Write-PlannerLog "    Benutzer konnte nicht zugewiesen werden: $userName" "WARN"
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
                        if ($_.Value.previewPriority) {
                            $references[$url]["previewPriority"] = $_.Value.previewPriority
                        }
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

    return @{
        NewPlanId    = $newPlan.id
        TasksCreated = $taskMapping.Count
        BucketsCreated = $bucketMapping.Count
    }
}

#endregion

#region Hauptprogramm

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

# Validiere Import-Verzeichnis
$pathValidationError = $null
if (-not (Test-SafePath -Path $ImportPath -Mode Import -ErrorMessage ([ref]$pathValidationError))) {
    Write-Host ""
    Write-Host "Fehler: $pathValidationError" -ForegroundColor Red
    Write-Host ""
    exit 1
}
Write-PlannerLog "Import-Verzeichnis: $ImportPath"

# Lade Index-Datei
$indexPath = Join-Path $ImportPath "_ExportIndex.json"
if (Test-Path $indexPath) {
    $indexData = Get-Content $indexPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-PlannerLog "Export-Index geladen. Export vom: $($indexData.ExportDate)"
    Write-PlannerLog "Exportiert von: $($indexData.ExportedBy)"
    Write-PlannerLog "Pläne im Export: $($indexData.TotalPlans)"
}

# JSON-Dateien finden
$jsonFiles = Get-ChildItem -Path $ImportPath -Filter "*.json" |
    Where-Object { $_.Name -ne "_ExportIndex.json" -and $_.Name -notmatch "ImportMapping" }

if ($jsonFiles.Count -eq 0) {
    Write-PlannerLog "Keine Export-Dateien gefunden in: $ImportPath" "ERROR"
    exit 1
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

# Microsoft Graph Module und Verbindung
Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

if (-not $DryRun) {
    if (-not (Connect-ToGraph)) {
        Write-PlannerLog "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
        exit 1
    }
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
}
else {
    # Success summary
    $totalTasksImported = ($importResults | Measure-Object -Property TasksCreated -Sum).Sum
    $totalBucketsImported = ($importResults | Measure-Object -Property BucketsCreated -Sum).Sum
    Write-Host "  Erfolgreich importiert:" -ForegroundColor White
    Write-Host "    Pläne:    $($importResults.Count)" -ForegroundColor Green
    Write-Host "    Buckets:  $totalBucketsImported" -ForegroundColor Green
    Write-Host "    Tasks:    $totalTasksImported" -ForegroundColor Green
    Write-Host ""

    # Cache statistics
    Write-CacheStatistics
}

# Error summary and exit code
$exitCode = Write-ErrorSummary -OutputPath $ImportPath

Write-PlannerLog "Import abgeschlossen mit Exit-Code: $exitCode"
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

exit $exitCode

#endregion
