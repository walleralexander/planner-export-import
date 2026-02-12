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
    [string]$ImportPath,

    [Parameter(Mandatory = $false)]
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
    [int]$ThrottleDelayMs = 500
)

#region Funktionen

function Write-Log {
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
    $logEntry | Out-File -FilePath "$ImportPath\import.log" -Append -Encoding UTF8
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
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Message -match "429") {
                # Rate Limited - warten und erneut versuchen
                $retryAfter = 30
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                Write-Log "Rate Limited. Warte $retryAfter Sekunden... (Versuch $attempt/$MaxRetries)" "WARN"
                Start-Sleep -Seconds $retryAfter
            }
            elseif ($attempt -ge $MaxRetries) {
                throw $_
            }
            else {
                Write-Log "Fehler bei Graph-Request (Versuch $attempt/$MaxRetries): $_" "WARN"
                Start-Sleep -Seconds (2 * $attempt)
            }
        }
    }
}

function Connect-ToGraph {
    Write-Log "Verbinde mit Microsoft Graph..."
    try {
        $context = Get-MgContext
        if ($null -eq $context) {
            try {
                # Versuche zuerst interaktive Anmeldung
                Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -NoWelcome -ErrorAction Stop
            }
            catch {
                # Fallback auf Device Code Flow wenn Browser-Auth fehlschlägt
                Write-Log "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." "WARN"
                Connect-MgGraph -Scopes "Group.ReadWrite.All", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -UseDeviceCode -NoWelcome
            }
        }
        $context = Get-MgContext
        if ($null -eq $context -or [string]::IsNullOrEmpty($context.Account)) {
            throw "Keine gültige Verbindung hergestellt"
        }
        Write-Log "Verbunden als: $($context.Account)" "OK"
        return $true
    }
    catch {
        Write-Log "Fehler bei der Verbindung: $_" "ERROR"
        return $false
    }
}

function Resolve-UserId {
    param([string]$OldUserId, [hashtable]$OldUserMap)

    # Wenn UserMapping vorhanden, verwende es
    if ($UserMapping -and $UserMapping.ContainsKey($OldUserId)) {
        return $UserMapping[$OldUserId]
    }

    # Versuche den User über UPN oder Mail in der neuen Umgebung zu finden
    if ($OldUserMap -and $OldUserMap[$OldUserId]) {
        $upn = $OldUserMap[$OldUserId].UserPrincipalName
        $mail = $OldUserMap[$OldUserId].Mail

        if ($upn) {
            try {
                $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                if ($user) {
                    return $user.id
                }
            }
            catch { }
        }

        if ($mail -and $mail -ne $upn) {
            try {
                $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mail`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                if ($user) {
                    return $user.id
                }
            }
            catch { }
        }
    }

    # Fallback: Versuche die alte ID direkt (gleicher Tenant)
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$OldUserId`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
        if ($user) {
            return $user.id
        }
    }
    catch { }

    return $null
}

function Import-PlanFromJson {
    param(
        [string]$JsonFilePath,
        [string]$TargetGroupId
    )

    Write-Log "Lade Export-Datei: $JsonFilePath"
    $planData = Get-Content $JsonFilePath -Raw -Encoding UTF8 | ConvertFrom-Json

    $planTitle = $planData.Plan.title
    $originalGroupId = $planData.Plan.groupId

    # Zielgruppe bestimmen
    $groupId = if ($TargetGroupId) { $TargetGroupId } else { $originalGroupId }

    if (-not $groupId) {
        Write-Log "Keine Zielgruppe angegeben und keine Original-Gruppe gefunden!" "ERROR"
        return $null
    }

    Write-Log "Erstelle Plan '$planTitle' in Gruppe $groupId..."

    if ($DryRun) {
        Write-Log "[DRY RUN] Würde Plan '$planTitle' erstellen" "DRYRUN"
        Write-Log "[DRY RUN] Buckets: $($planData.Buckets.Count)" "DRYRUN"
        Write-Log "[DRY RUN] Tasks: $($planData.Tasks.Count)" "DRYRUN"
        return
    }

    # 1. Plan erstellen
    $newPlan = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/plans" -Body @{
        owner = $groupId
        title = $planTitle
    }
    Write-Log "  Plan erstellt: $($newPlan.id)" "OK"

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
            Write-Log "  Kategorien gesetzt" "OK"
        }
        catch {
            Write-Log "  Fehler beim Setzen der Kategorien: $_" "WARN"
        }
    }

    # 3. Buckets erstellen (ID-Mapping für Tasks)
    $bucketMapping = @{}
    foreach ($bucket in ($planData.Buckets | Sort-Object orderHint)) {
        try {
            $newBucket = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/buckets" -Body @{
                name    = $bucket.name
                planId  = $newPlan.id
            }
            $bucketMapping[$bucket.id] = $newBucket.id
            Write-Log "  Bucket erstellt: $($bucket.name)" "OK"
        }
        catch {
            Write-Log "  Fehler beim Erstellen von Bucket '$($bucket.name)': $_" "ERROR"
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
        Write-Progress -Activity "Importiere Tasks für '$planTitle'" -Status "Task $taskCounter von $totalTasks: $($task.title)" -PercentComplete (($taskCounter / $totalTasks) * 100)

        # Abgeschlossene Tasks überspringen wenn gewünscht
        if ($SkipCompletedTasks -and $task.percentComplete -eq 100) {
            Write-Log "  Task übersprungen (abgeschlossen): $($task.title)"
            continue
        }

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
                    Write-Log "    Benutzer konnte nicht zugewiesen werden: $userName" "WARN"
                }
            }
            if ($assignments.Count -gt 0) {
                $taskBody["assignments"] = $assignments
            }
        }

        try {
            $newTask = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/tasks" -Body $taskBody
            $taskMapping[$task.id] = $newTask.id
            Write-Log "  Task erstellt: $($task.title)" "OK"

            # 5. Task-Details setzen (Beschreibung, Checkliste, Referenzen)
            $detail = $planData.TaskDetails | Where-Object { $_.taskId -eq $task.id }
            if ($detail) {
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
                        Write-Log "    Details gesetzt für: $($task.title)" "OK"
                    }
                    catch {
                        Write-Log "    Fehler beim Setzen der Task-Details: $_" "WARN"
                    }
                }
            }
        }
        catch {
            Write-Log "  Fehler beim Erstellen von Task '$($task.title)': $_" "ERROR"
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
    $mappingData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$ImportPath\${planFileName}_ImportMapping.json" -Encoding UTF8

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

# Prüfe Import-Verzeichnis
if (-not (Test-Path $ImportPath)) {
    Write-Log "Import-Verzeichnis nicht gefunden: $ImportPath" "ERROR"
    exit 1
}

# Lade Index-Datei
$indexPath = Join-Path $ImportPath "_ExportIndex.json"
if (Test-Path $indexPath) {
    $indexData = Get-Content $indexPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Log "Export-Index geladen. Export vom: $($indexData.ExportDate)"
    Write-Log "Exportiert von: $($indexData.ExportedBy)"
    Write-Log "Pläne im Export: $($indexData.TotalPlans)"
}

# JSON-Dateien finden
$jsonFiles = Get-ChildItem -Path $ImportPath -Filter "*.json" | 
    Where-Object { $_.Name -ne "_ExportIndex.json" -and $_.Name -notmatch "ImportMapping" }

if ($jsonFiles.Count -eq 0) {
    Write-Log "Keine Export-Dateien gefunden in: $ImportPath" "ERROR"
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
        Write-Log "Import abgebrochen durch Benutzer." "WARN"
        exit 0
    }
}

# Microsoft Graph Module und Verbindung
Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

if (-not $DryRun) {
    if (-not (Connect-ToGraph)) {
        Write-Log "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
        exit 1
    }
}

# Import durchführen
$importResults = @()
foreach ($jsonFile in $jsonFiles) {
    Write-Host ""
    Write-Log "=== Importiere: $($jsonFile.Name) ===" "OK"
    
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
}
else {
    $totalTasksImported = ($importResults | Measure-Object -Property TasksCreated -Sum).Sum
    $totalBucketsImported = ($importResults | Measure-Object -Property BucketsCreated -Sum).Sum
    Write-Host "  Pläne importiert:  $($importResults.Count)" -ForegroundColor White
    Write-Host "  Buckets erstellt:  $totalBucketsImported" -ForegroundColor White
    Write-Host "  Tasks erstellt:    $totalTasksImported" -ForegroundColor White
}

Write-Host ""
Write-Log "Import abgeschlossen."

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

#endregion
