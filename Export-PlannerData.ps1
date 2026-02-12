<#
.SYNOPSIS
    Exportiert alle Microsoft Planner-Daten (Pläne, Buckets, Tasks, Details) via Microsoft Graph API.

.DESCRIPTION
    Dieses Script liest alle Planner-Pläne aus, die dem angemeldeten Benutzer bzw. den
    angegebenen Microsoft 365 Gruppen zugeordnet sind, und exportiert sämtliche Daten
    (Buckets, Tasks inkl. Checklisten, Anhänge, Beschreibungen, Zuweisungen etc.)
    in JSON-Dateien, die für einen späteren Re-Import verwendet werden können.

.NOTES
    Voraussetzungen:
    - PowerShell 5.1 oder höher (empfohlen: PowerShell 7+)
    - Microsoft.Graph PowerShell Module
      Install-Module Microsoft.Graph -Scope CurrentUser
    - Berechtigungen: Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read

.AUTHOR
    Alexander Waller
    Datum: 2026-02-09
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\planner-data\PlannerExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [string[]]$GroupIds,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCompletedTasks
)

#region Funktionen

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $(switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "OK"    { "Green" }
        default { "White" }
    })
    $logEntry | Out-File -FilePath "$ExportPath\export.log" -Append -Encoding UTF8
}

function Connect-ToGraph {
    Write-Log "Verbinde mit Microsoft Graph..."
    try {
        # Prüfe ob bereits verbunden
        $context = Get-MgContext
        if ($null -eq $context) {
            try {
                # Versuche zuerst interaktive Anmeldung
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -NoWelcome -ErrorAction Stop
            }
            catch {
                # Fallback auf Device Code Flow wenn Browser-Auth fehlschlägt
                Write-Log "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." "WARN"
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -UseDeviceCode -NoWelcome
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
        Write-Log "Fehler bei der Verbindung zu Microsoft Graph: $_" "ERROR"
        return $false
    }
}

function Get-AllUserPlans {
    Write-Log "Lade alle Pläne des Benutzers..."
    $plans = @()

    try {
        # Methode 1: Über die Gruppen des Benutzers
        $myGroups = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?`$filter=groupTypes/any(g:g eq 'Unified')&`$select=id,displayName" -OutputType PSObject
        
        if ($myGroups.value) {
            foreach ($group in $myGroups.value) {
                Write-Log "Prüfe Gruppe: $($group.displayName) ($($group.id))"
                try {
                    $groupPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/planner/plans" -OutputType PSObject
                    if ($groupPlans.value) {
                        foreach ($plan in $groupPlans.value) {
                            $plan | Add-Member -NotePropertyName "groupDisplayName" -NotePropertyValue $group.displayName -Force
                            $plan | Add-Member -NotePropertyName "groupId" -NotePropertyValue $group.id -Force
                            $plans += $plan
                            Write-Log "  Plan gefunden: $($plan.title)" "OK"
                        }
                    }
                }
                catch {
                    Write-Log "  Keine Pläne in Gruppe $($group.displayName) oder kein Zugriff" "WARN"
                }
            }
        }

        # Methode 2: Direkt über /me/planner/plans (falls unterstützt)
        try {
            $myPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/planner/plans" -OutputType PSObject
            if ($myPlans.value) {
                foreach ($plan in $myPlans.value) {
                    if ($plans.id -notcontains $plan.id) {
                        $plans += $plan
                        Write-Log "  Zusätzlicher Plan gefunden: $($plan.title)" "OK"
                    }
                }
            }
        }
        catch {
            Write-Log "Direkter Planzugriff nicht verfügbar, nutze Gruppen-Methode" "WARN"
        }
    }
    catch {
        Write-Log "Fehler beim Laden der Pläne: $_" "ERROR"
    }

    Write-Log "Insgesamt $($plans.Count) Pläne gefunden"
    return $plans
}

function Get-PlansByGroupIds {
    param([string[]]$GroupIds)
    $plans = @()

    foreach ($groupId in $GroupIds) {
        Write-Log "Lade Pläne für Gruppe: $groupId"
        try {
            # Gruppenname laden
            $group = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,displayName" -OutputType PSObject

            $groupPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/planner/plans" -OutputType PSObject
            if ($groupPlans.value) {
                foreach ($plan in $groupPlans.value) {
                    $plan | Add-Member -NotePropertyName "groupDisplayName" -NotePropertyValue $group.displayName -Force
                    $plan | Add-Member -NotePropertyName "groupId" -NotePropertyValue $groupId -Force
                    $plans += $plan
                    Write-Log "  Plan gefunden: $($plan.title)" "OK"
                }
            }
        }
        catch {
            Write-Log "Fehler beim Laden der Pläne für Gruppe $groupId : $_" "ERROR"
        }
    }

    return $plans
}

function Export-PlanDetails {
    param(
        [PSObject]$Plan,
        [string]$PlanExportPath
    )

    $planData = @{
        Plan       = $Plan
        Buckets    = @()
        Tasks      = @()
        TaskDetails = @()
        Categories = @{}
    }

    # Plan-Details laden (Kategorien/Labels)
    try {
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/details" -OutputType PSObject
        $planData.Categories = $planDetails.categoryDescriptions
        $planData["PlanDetails"] = $planDetails

        # ETag speichern für späteren Import
        if ($planDetails.'@odata.etag') {
            $planData["PlanDetailsEtag"] = $planDetails.'@odata.etag'
        }
    }
    catch {
        Write-Log "  Konnte Plan-Details nicht laden: $_" "WARN"
    }

    # Buckets laden
    try {
        $buckets = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/buckets" -OutputType PSObject
        $planData.Buckets = $buckets.value
        Write-Log "  $($buckets.value.Count) Buckets geladen"
    }
    catch {
        Write-Log "  Fehler beim Laden der Buckets: $_" "ERROR"
    }

    # Tasks laden (mit Paging)
    try {
        $allTasks = @()
        $tasksUri = "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/tasks"
        
        do {
            $tasksResponse = Invoke-MgGraphRequest -Method GET -Uri $tasksUri -OutputType PSObject
            if ($tasksResponse.value) {
                $allTasks += $tasksResponse.value
            }
            $tasksUri = $tasksResponse.'@odata.nextLink'
        } while ($tasksUri)

        # Optional: Abgeschlossene Tasks filtern
        if (-not $IncludeCompletedTasks) {
            $completedCount = ($allTasks | Where-Object { $_.percentComplete -eq 100 }).Count
            Write-Log "  $completedCount abgeschlossene Tasks werden übersprungen (verwende -IncludeCompletedTasks um diese einzubeziehen)"
            $allTasks = $allTasks | Where-Object { $_.percentComplete -ne 100 }
        }

        $planData.Tasks = $allTasks
        Write-Log "  $($allTasks.Count) Tasks geladen"

        # Task-Details laden (Beschreibungen, Checklisten, Referenzen)
        $taskDetails = @()
        $counter = 0
        foreach ($task in $allTasks) {
            $counter++
            Write-Progress -Activity "Lade Task-Details für '$($Plan.title)'" -Status "Task $counter von $($allTasks.Count)" -PercentComplete (($counter / $allTasks.Count) * 100)
            
            try {
                $detail = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details" -OutputType PSObject
                $detail | Add-Member -NotePropertyName "taskId" -NotePropertyValue $task.id -Force
                $taskDetails += $detail

                # Rate Limiting vermeiden
                Start-Sleep -Milliseconds 200
            }
            catch {
                Write-Log "    Fehler bei Task-Detail $($task.title): $_" "WARN"
            }
        }
        Write-Progress -Activity "Lade Task-Details" -Completed
        $planData.TaskDetails = $taskDetails
        Write-Log "  $($taskDetails.Count) Task-Details geladen"
    }
    catch {
        Write-Log "  Fehler beim Laden der Tasks: $_" "ERROR"
    }

    # Benutzerinfo für Zuweisungen auflösen
    $userIds = @()
    foreach ($task in $planData.Tasks) {
        if ($task.assignments) {
            $task.assignments.PSObject.Properties | ForEach-Object {
                if ($_.Name -notin $userIds -and $_.Name -match '^[0-9a-f-]+$') {
                    $userIds += $_.Name
                }
            }
        }
        if ($task.createdBy.user.id -and $task.createdBy.user.id -notin $userIds) {
            $userIds += $task.createdBy.user.id
        }
    }

    $userMap = @{}
    foreach ($userId in $userIds | Select-Object -Unique) {
        try {
            $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$userId`?`$select=id,displayName,userPrincipalName,mail" -OutputType PSObject
            $userMap[$userId] = @{
                Id                = $user.id
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                Mail              = $user.mail
            }
        }
        catch {
            Write-Log "    Konnte Benutzer $userId nicht auflösen" "WARN"
            $userMap[$userId] = @{ Id = $userId; DisplayName = "Unbekannt"; UserPrincipalName = ""; Mail = "" }
        }
    }
    $planData["UserMap"] = $userMap

    # Exportieren
    $planFileName = ($Plan.title -replace '[\\/:*?"<>|]', '_')
    $planFilePath = Join-Path $PlanExportPath "$planFileName.json"

    $planData | ConvertTo-Json -Depth 20 | Out-File -FilePath $planFilePath -Encoding UTF8
    Write-Log "  Plan exportiert nach: $planFilePath" "OK"

    # Zusätzlich eine lesbare Zusammenfassung erstellen
    Export-ReadableSummary -PlanData $planData -OutputPath (Join-Path $PlanExportPath "$planFileName`_Zusammenfassung.txt")

    return $planData
}

function Export-ReadableSummary {
    param(
        [hashtable]$PlanData,
        [string]$OutputPath
    )

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine("PLANNER EXPORT - ZUSAMMENFASSUNG")
    [void]$sb.AppendLine("Plan: $($PlanData.Plan.title)")
    if ($PlanData.Plan.groupDisplayName) {
        [void]$sb.AppendLine("Gruppe: $($PlanData.Plan.groupDisplayName)")
    }
    [void]$sb.AppendLine("Exportiert am: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')")
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine("")

    # Kategorien/Labels
    if ($PlanData.Categories) {
        [void]$sb.AppendLine("KATEGORIEN / LABELS:")
        $PlanData.Categories.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object {
            [void]$sb.AppendLine("  $($_.Name): $($_.Value)")
        }
        [void]$sb.AppendLine("")
    }

    # Buckets und Tasks
    foreach ($bucket in ($PlanData.Buckets | Sort-Object orderHint)) {
        [void]$sb.AppendLine("-" * 60)
        [void]$sb.AppendLine("BUCKET: $($bucket.name)")
        [void]$sb.AppendLine("-" * 60)

        $bucketTasks = $PlanData.Tasks | Where-Object { $_.bucketId -eq $bucket.id } | Sort-Object orderHint
        if ($bucketTasks.Count -eq 0) {
            [void]$sb.AppendLine("  (keine Tasks)")
        }
        else {
            foreach ($task in $bucketTasks) {
                $status = switch ($task.percentComplete) {
                    0   { "[ ] Nicht begonnen" }
                    50  { "[~] In Bearbeitung" }
                    100 { "[x] Abgeschlossen" }
                    default { "[$($task.percentComplete)%]" }
                }
                
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("  $status $($task.title)")

                # Priorität
                $priority = switch ($task.priority) {
                    0 { "Dringend" }
                    1 { "Wichtig" }
                    2 { "Mittel" }
                    default { "Niedrig" }
                }
                [void]$sb.AppendLine("    Priorität: $priority")

                # Fälligkeitsdatum
                if ($task.dueDateTime) {
                    $dueDate = [DateTime]::Parse($task.dueDateTime).ToString("dd.MM.yyyy")
                    [void]$sb.AppendLine("    Fällig am: $dueDate")
                }

                # Startdatum
                if ($task.startDateTime) {
                    $startDate = [DateTime]::Parse($task.startDateTime).ToString("dd.MM.yyyy")
                    [void]$sb.AppendLine("    Start: $startDate")
                }

                # Zuweisungen
                if ($task.assignments) {
                    $assignees = @()
                    $task.assignments.PSObject.Properties | ForEach-Object {
                        if ($PlanData.UserMap[$_.Name]) {
                            $assignees += $PlanData.UserMap[$_.Name].DisplayName
                        }
                    }
                    if ($assignees.Count -gt 0) {
                        [void]$sb.AppendLine("    Zugewiesen an: $($assignees -join ', ')")
                    }
                }

                # Labels/Kategorien
                if ($task.appliedCategories) {
                    $labels = @()
                    $task.appliedCategories.PSObject.Properties | Where-Object { $_.Value -eq $true } | ForEach-Object {
                        if ($PlanData.Categories -and $PlanData.Categories.$($_.Name)) {
                            $labels += $PlanData.Categories.$($_.Name)
                        }
                        else {
                            $labels += $_.Name
                        }
                    }
                    if ($labels.Count -gt 0) {
                        [void]$sb.AppendLine("    Labels: $($labels -join ', ')")
                    }
                }

                # Task-Details
                $detail = $PlanData.TaskDetails | Where-Object { $_.taskId -eq $task.id }
                if ($detail) {
                    # Beschreibung
                    if ($detail.description) {
                        [void]$sb.AppendLine("    Beschreibung:")
                        $detail.description -split "`n" | ForEach-Object {
                            [void]$sb.AppendLine("      $_")
                        }
                    }

                    # Checkliste
                    if ($detail.checklist) {
                        [void]$sb.AppendLine("    Checkliste:")
                        $detail.checklist.PSObject.Properties | ForEach-Object {
                            $checkItem = $_.Value
                            $checkMark = if ($checkItem.isChecked) { "[x]" } else { "[ ]" }
                            [void]$sb.AppendLine("      $checkMark $($checkItem.title)")
                        }
                    }

                    # Referenzen/Links
                    if ($detail.references) {
                        $refs = $detail.references.PSObject.Properties
                        if ($refs.Count -gt 0) {
                            [void]$sb.AppendLine("    Anhänge/Links:")
                            $refs | ForEach-Object {
                                $ref = $_.Value
                                $url = $_.Name -replace '%2F', '/' -replace '%3A', ':'
                                [void]$sb.AppendLine("      - $($ref.alias): $url")
                            }
                        }
                    }
                }
            }
        }
        [void]$sb.AppendLine("")
    }

    # Tasks ohne Bucket
    $orphanTasks = $PlanData.Tasks | Where-Object { -not $_.bucketId -or ($PlanData.Buckets.id -notcontains $_.bucketId) }
    if ($orphanTasks.Count -gt 0) {
        [void]$sb.AppendLine("-" * 60)
        [void]$sb.AppendLine("TASKS OHNE BUCKET:")
        [void]$sb.AppendLine("-" * 60)
        foreach ($task in $orphanTasks) {
            [void]$sb.AppendLine("  - $($task.title)")
        }
    }

    # Statistik
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine("STATISTIK:")
    [void]$sb.AppendLine("  Buckets: $($PlanData.Buckets.Count)")
    [void]$sb.AppendLine("  Tasks gesamt: $($PlanData.Tasks.Count)")
    [void]$sb.AppendLine("  Nicht begonnen: $(($PlanData.Tasks | Where-Object { $_.percentComplete -eq 0 }).Count)")
    [void]$sb.AppendLine("  In Bearbeitung: $(($PlanData.Tasks | Where-Object { $_.percentComplete -eq 50 }).Count)")
    [void]$sb.AppendLine("  Abgeschlossen: $(($PlanData.Tasks | Where-Object { $_.percentComplete -eq 100 }).Count)")
    [void]$sb.AppendLine("=" * 80)

    $sb.ToString() | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Log "  Zusammenfassung erstellt: $OutputPath" "OK"
}

#endregion

#region Hauptprogramm

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Planner Export Tool" -ForegroundColor Cyan
Write-Host "  by Alexander Waller" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Export-Verzeichnis erstellen
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}
Write-Log "Export-Verzeichnis: $ExportPath"

# Microsoft.Graph Modul prüfen
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Planner)) {
    Write-Log "Microsoft.Graph Module werden installiert..." "WARN"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

# Verbinden
if (-not (Connect-ToGraph)) {
    Write-Log "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
    exit 1
}

# Pläne laden
$plans = @()
if ($GroupIds) {
    $plans = Get-PlansByGroupIds -GroupIds $GroupIds
}
else {
    $plans = Get-AllUserPlans
}

if ($plans.Count -eq 0) {
    Write-Log "Keine Pläne gefunden. Überprüfen Sie die Berechtigungen." "ERROR"
    exit 1
}

# Übersicht anzeigen
Write-Host ""
Write-Host "Gefundene Pläne:" -ForegroundColor Yellow
for ($i = 0; $i -lt $plans.Count; $i++) {
    $groupName = if ($plans[$i].groupDisplayName) { " (Gruppe: $($plans[$i].groupDisplayName))" } else { "" }
    Write-Host "  [$($i+1)] $($plans[$i].title)$groupName"
}
Write-Host ""

# Alle Pläne exportieren
$exportSummary = @()
foreach ($plan in $plans) {
    Write-Host ""
    Write-Log "=== Exportiere Plan: $($plan.title) ===" "OK"
    $planData = Export-PlanDetails -Plan $plan -PlanExportPath $ExportPath
    $exportSummary += @{
        PlanTitle  = $plan.title
        GroupName  = $plan.groupDisplayName
        Buckets    = $planData.Buckets.Count
        Tasks      = $planData.Tasks.Count
    }
}

# Gesamtübersicht als Index-Datei
$indexData = @{
    ExportDate    = (Get-Date).ToString("o")
    ExportedBy    = (Get-MgContext).Account
    TotalPlans    = $plans.Count
    Plans         = $exportSummary
    ScriptVersion = "1.0.0"
}
$indexData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$ExportPath\_ExportIndex.json" -Encoding UTF8

# Zusammenfassung
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  EXPORT ABGESCHLOSSEN" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Verzeichnis: $ExportPath" -ForegroundColor White
Write-Host "  Pläne:       $($plans.Count)" -ForegroundColor White
$totalTasks = ($exportSummary | Measure-Object -Property Tasks -Sum).Sum
Write-Host "  Tasks:       $totalTasks" -ForegroundColor White
Write-Host ""
Write-Host "  Dateien pro Plan:" -ForegroundColor Yellow
Write-Host "    - <PlanName>.json         (Strukturierte Daten für Import)" -ForegroundColor Gray
Write-Host "    - <PlanName>_Zusammenfassung.txt  (Lesbare Übersicht)" -ForegroundColor Gray
Write-Host "    - _ExportIndex.json       (Gesamtübersicht)" -ForegroundColor Gray
Write-Host ""
Write-Log "Export abgeschlossen. $($plans.Count) Pläne mit $totalTasks Tasks exportiert." "OK"

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

#endregion
