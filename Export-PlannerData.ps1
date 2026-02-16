<#
.SYNOPSIS
    Exportiert Microsoft Planner-Daten (Pläne, Buckets, Tasks, Details) aus M365-Gruppen via Microsoft Graph API.

.DESCRIPTION
    Dieses Script liest Planner-Pläne aus Microsoft 365 und exportiert sämtliche Daten
    (Buckets, Tasks inkl. Checklisten, Anhänge, Beschreibungen, Zuweisungen etc.)
    in JSON-Dateien, die für einen späteren Re-Import verwendet werden können.

    Zwei Export-Modi verfügbar:
    1. User-basiert: Alle Pläne des aktuellen Benutzers (-UseCurrentUser)
    2. Gruppen-basiert: Pläne aus spezifischen M365-Gruppen (-GroupNames/-GroupIds/-Interactive)

.PARAMETER ExportPath
    Zielpfad für den Export (Standard: C:\planner-data\PlannerExport_YYYYMMDD_HHMMSS)

.PARAMETER GroupNames
    Namen der M365-Gruppen/SharePoint-Seiten, aus denen Pläne exportiert werden sollen.
    Beispiel: -GroupNames "Projektteam Alpha", "Marketing"

.PARAMETER GroupIds
    Direkte Angabe von M365-Gruppen-IDs (alternative zu GroupNames).
    Beispiel: -GroupIds "abc-123-def", "xyz-789-uvw"

.PARAMETER Interactive
    Zeigt eine interaktive Auswahl aller verfügbaren M365-Gruppen.

.PARAMETER IncludeCompletedTasks
    Exportiert auch abgeschlossene Tasks (Standard: nur aktive Tasks).

.PARAMETER UseCurrentUser
    Exportiert alle Pläne des aktuellen Benutzers aus allen seinen M365-Gruppen.
    Dies ist eine Alternative zu -GroupIds, -GroupNames oder -Interactive.

.EXAMPLE
    .\Export-PlannerData.ps1 -UseCurrentUser
    Exportiert alle Pläne des angemeldeten Benutzers aus allen seinen Gruppen

.EXAMPLE
    .\Export-PlannerData.ps1 -UseCurrentUser -IncludeCompletedTasks
    Exportiert alle Pläne des Users inkl. abgeschlossener Tasks

.EXAMPLE
    .\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha"
    Exportiert alle Pläne der Gruppe "Projektteam Alpha"

.EXAMPLE
    .\Export-PlannerData.ps1 -Interactive
    Zeigt interaktive Gruppenauswahl

.EXAMPLE
    .\Export-PlannerData.ps1 -GroupIds "abc-123-def" -IncludeCompletedTasks
    Exportiert Pläne der angegebenen Gruppe inkl. abgeschlossener Tasks

.NOTES
    Voraussetzungen:
    - PowerShell 5.1 oder höher (empfohlen: PowerShell 7+)
    - Microsoft.Graph PowerShell Module
      Install-Module Microsoft.Graph -Scope CurrentUser
    - Berechtigungen: Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read

.AUTHOR
    Alexander Waller
    Datum: 2026-02-16
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\planner-data\PlannerExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [string[]]$GroupIds,

    [Parameter(Mandatory = $false)]
    [string[]]$GroupNames,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCompletedTasks,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$UseCurrentUser
)

#region Funktionen

function Write-PlannerLog {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $(switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "OK"    { "Green" }
        default { "White" }
    })
    try {
        $logEntry | Out-File -FilePath "$ExportPath\export.log" -Append -Encoding utf8BOM -ErrorAction Stop
    }
    catch {
        Write-Host "[ERROR] Konnte nicht in Log-Datei schreiben: $_" -ForegroundColor Red
    }
}

function Connect-ToGraph {
    Write-PlannerLog "Verbinde mit Microsoft Graph..."
    try {
        # PrÃ¼fe ob bereits verbunden
        $context = Get-MgContext
        if ($null -eq $context) {
            try {
                # Versuche zuerst interaktive Anmeldung
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -NoWelcome -ErrorAction Stop
            }
            catch {
                # Fallback auf Device Code Flow wenn Browser-Auth fehlschlÃ¤gt
                Write-PlannerLog "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." "WARN"
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -UseDeviceCode -NoWelcome
            }
        }
        $context = Get-MgContext
        if ($null -eq $context -or [string]::IsNullOrEmpty($context.Account)) {
            throw "Keine gÃ¼ltige Verbindung hergestellt"
        }
        Write-PlannerLog "Verbunden als: $($context.Account)" "OK"
        return $true
    }
    catch {
        Write-PlannerLog "Fehler bei der Verbindung zu Microsoft Graph: $_" "ERROR"
        return $false
    }
}

function Get-AllM365Groups {
    Write-PlannerLog "Lade alle verfügbaren M365-Gruppen..."
    $groups = @()

    try {
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(g:g eq 'Unified')&`$select=id,displayName,mail&`$orderby=displayName"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
            if ($response.value) {
                $groups += $response.value
            }
            $uri = $response.'@odata.nextLink'
        } while ($uri)

        Write-PlannerLog "$($groups.Count) M365-Gruppen gefunden"
        return $groups
    }
    catch {
        Write-PlannerLog "Fehler beim Laden der M365-Gruppen: $_" "ERROR"
        return @()
    }
}

function Get-GroupsByNames {
    param([string[]]$GroupNames)

    if ($null -eq $GroupNames -or $GroupNames.Count -eq 0) {
        Write-PlannerLog "Keine Gruppennamen angegeben" "ERROR"
        return @()
    }

    Write-PlannerLog "Suche Gruppen nach Namen: $($GroupNames -join ', ')"
    $foundGroups = @()

    foreach ($name in $GroupNames) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            Write-PlannerLog "  Überspringe leeren Gruppennamen" "WARN"
            continue
        }

        try {
            # Suche nach exaktem Namen oder ähnlichem Namen
            # Hinweis: Sonderzeichen in Gruppennamen könnten zu Problemen führen
            $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(g:g eq 'Unified') and (displayName eq '$name' or startswith(displayName,'$name'))&`$select=id,displayName,mail"

            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

            if ($null -eq $response) {
                Write-PlannerLog "  Keine Antwort vom Server für Gruppe: $name" "WARN"
            }
            elseif ($response.value -and $response.value.Count -gt 0) {
                $matchCount = $response.value.Count
                Write-PlannerLog "  $matchCount Gruppe(n) gefunden für: $name"
                foreach ($group in $response.value) {
                    # Duplikate vermeiden
                    if ($foundGroups.id -notcontains $group.id) {
                        $foundGroups += $group
                        Write-PlannerLog "    -> $($group.displayName) ($($group.id))" "OK"
                    }
                }
            }
            else {
                Write-PlannerLog "  Keine Gruppe gefunden mit Namen: $name" "WARN"
            }
        }
        catch {
            Write-PlannerLog "  Fehler bei der Suche nach Gruppe '$name': $_" "ERROR"
        }
    }

    Write-PlannerLog "Insgesamt $($foundGroups.Count) eindeutige Gruppe(n) gefunden"
    return $foundGroups
}

function Show-GroupSelectionMenu {
    param([array]$Groups)

    Write-Host ""
    Write-Host "Verfügbare M365-Gruppen:" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $Groups.Count; $i++) {
        Write-Host "  [$($i+1)] $($Groups[$i].displayName)" -ForegroundColor White
        if ($Groups[$i].mail) {
            Write-Host "      $($Groups[$i].mail)" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "  [A] Alle Gruppen" -ForegroundColor Cyan
    Write-Host "  [0] Abbrechen" -ForegroundColor Red
    Write-Host ""

    $selection = Read-Host "Bitte wählen Sie eine oder mehrere Gruppen (z.B. 1,3,5 oder A)"

    if ($selection -eq "0") {
        return @()
    }
    elseif ($selection -eq "A" -or $selection -eq "a") {
        return $Groups
    }
    else {
        $selectedGroups = @()
        $indices = $selection -split ',' | ForEach-Object { $_.Trim() }

        foreach ($index in $indices) {
            if ($index -match '^\d+$') {
                $idx = [int]$index - 1
                if ($idx -ge 0 -and $idx -lt $Groups.Count) {
                    $selectedGroups += $Groups[$idx]
                }
            }
        }

        return $selectedGroups
    }
}

function Get-AllUserPlans {
    Write-PlannerLog "Lade alle Pläne des Benutzers..."
    $plans = @()

    try {
        # Methode 1: Über die Gruppen des Benutzers
        $myGroups = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group?`$filter=groupTypes/any(g:g eq 'Unified')&`$select=id,displayName" -OutputType PSObject

        if ($null -eq $myGroups) {
            Write-PlannerLog "Keine Antwort beim Abrufen der Benutzergruppen" "WARN"
        }
        elseif ($myGroups.value -and $myGroups.value.Count -gt 0) {
            Write-PlannerLog "$($myGroups.value.Count) M365-Gruppen des Benutzers gefunden"
            foreach ($group in $myGroups.value) {
                Write-PlannerLog "Prüfe Gruppe: $($group.displayName) ($($group.id))"
                try {
                    $groupPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/planner/plans" -OutputType PSObject
                    if ($groupPlans.value -and $groupPlans.value.Count -gt 0) {
                        foreach ($plan in $groupPlans.value) {
                            $plan | Add-Member -NotePropertyName "groupDisplayName" -NotePropertyValue $group.displayName -Force
                            $plan | Add-Member -NotePropertyName "groupId" -NotePropertyValue $group.id -Force
                            $plans += $plan
                            Write-PlannerLog "  Plan gefunden: $($plan.title)" "OK"
                        }
                    }
                }
                catch {
                    Write-PlannerLog "  Keine Pläne in Gruppe $($group.displayName) oder kein Zugriff: $_" "WARN"
                }
            }
        }
        else {
            Write-PlannerLog "Benutzer ist kein Mitglied von M365-Gruppen" "WARN"
        }

        # Methode 2: Direkt über /me/planner/plans (falls unterstützt)
        try {
            $myPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/planner/plans" -OutputType PSObject
            if ($myPlans.value -and $myPlans.value.Count -gt 0) {
                foreach ($plan in $myPlans.value) {
                    if ($plans.id -notcontains $plan.id) {
                        $plans += $plan
                        Write-PlannerLog "  Zusätzlicher Plan gefunden: $($plan.title)" "OK"
                    }
                }
            }
        }
        catch {
            Write-PlannerLog "Direkter Planzugriff nicht verfügbar, nutze Gruppen-Methode: $_" "WARN"
        }
    }
    catch {
        Write-PlannerLog "Fehler beim Laden der Pläne: $_" "ERROR"
        return @()
    }

    Write-PlannerLog "Insgesamt $($plans.Count) Pläne gefunden"
    return $plans
}

function Get-PlansByGroupIds {
    param([string[]]$GroupIds)
    $plans = @()

    foreach ($groupId in $GroupIds) {
        Write-PlannerLog "Lade Pläne für Gruppe: $groupId"
        try {
            # Gruppenname laden
            $group = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId`?`$select=id,displayName" -OutputType PSObject

            $groupPlans = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/planner/plans" -OutputType PSObject
            if ($groupPlans.value) {
                foreach ($plan in $groupPlans.value) {
                    $plan | Add-Member -NotePropertyName "groupDisplayName" -NotePropertyValue $group.displayName -Force
                    $plan | Add-Member -NotePropertyName "groupId" -NotePropertyValue $groupId -Force
                    $plans += $plan
                    Write-PlannerLog "  Plan gefunden: $($plan.title)" "OK"
                }
            }
        }
        catch {
            Write-PlannerLog "Fehler beim Laden der Pläne für Gruppe $groupId : $_" "ERROR"
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
        Write-PlannerLog "  Konnte Plan-Details nicht laden: $_" "WARN"
    }

    # Buckets laden
    try {
        $buckets = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/buckets" -OutputType PSObject
        $planData.Buckets = $buckets.value
        Write-PlannerLog "  $($buckets.value.Count) Buckets geladen"
    }
    catch {
        Write-PlannerLog "  Fehler beim Laden der Buckets: $_" "ERROR"
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
            Write-PlannerLog "  $completedCount abgeschlossene Tasks werden übersprungen (verwende -IncludeCompletedTasks um diese einzubeziehen)"
            $allTasks = $allTasks | Where-Object { $_.percentComplete -ne 100 }
        }

        $planData.Tasks = $allTasks
        Write-PlannerLog "  $($allTasks.Count) Tasks geladen"

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
                Write-PlannerLog "    Fehler bei Task-Detail $($task.title): $_" "WARN"
            }
        }
        Write-Progress -Activity "Lade Task-Details" -Completed
        $planData.TaskDetails = $taskDetails
        Write-PlannerLog "  $($taskDetails.Count) Task-Details geladen"
    }
    catch {
        Write-PlannerLog "  Fehler beim Laden der Tasks: $_" "ERROR"
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
            Write-PlannerLog "    Konnte Benutzer $userId nicht auflÃ¶sen" "WARN"
            $userMap[$userId] = @{ Id = $userId; DisplayName = "Unbekannt"; UserPrincipalName = ""; Mail = "" }
        }
    }
    $planData["UserMap"] = $userMap

    # Exportieren
    $planFileName = ($Plan.title -replace '[\\/:*?"<>|]', '_')
    $planFilePath = Join-Path $PlanExportPath "$planFileName.json"

    try {
        $planData | ConvertTo-Json -Depth 20 | Out-File -FilePath $planFilePath -Encoding utf8BOM -ErrorAction Stop
    }
    catch {
        Write-PlannerLog "Fehler beim Schreiben der JSON-Datei: $planFilePath - $_" "ERROR"
        throw
    }
    Write-PlannerLog "  Plan exportiert nach: $planFilePath" "OK"

    # ZusÃ¤tzlich eine lesbare Zusammenfassung erstellen
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

                # PrioritÃ¤t
                $priority = switch ($task.priority) {
                    0 { "Dringend" }
                    1 { "Wichtig" }
                    2 { "Mittel" }
                    default { "Niedrig" }
                }
                [void]$sb.AppendLine("    PrioritÃ¤t: $priority")

                # FÃ¤lligkeitsdatum
                if ($task.dueDateTime) {
                    $dueDate = [DateTime]::Parse($task.dueDateTime).ToString("dd.MM.yyyy")
                    [void]$sb.AppendLine("    FÃ¤llig am: $dueDate")
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
                            [void]$sb.AppendLine("    AnhÃ¤nge/Links:")
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

    try {
        $sb.ToString() | Out-File -FilePath $OutputPath -Encoding utf8BOM -ErrorAction Stop
        Write-PlannerLog "  Zusammenfassung erstellt: $OutputPath" "OK"
    }
    catch {
        Write-PlannerLog "Fehler beim Schreiben der Zusammenfassung: $OutputPath - $_" "ERROR"
    }
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
Write-PlannerLog "Export-Verzeichnis: $ExportPath"

# Microsoft.Graph Modul prüfen
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Planner)) {
    Write-PlannerLog "Microsoft.Graph Module werden installiert..." "WARN"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

# Verbinden
if (-not (Connect-ToGraph)) {
    Write-PlannerLog "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
    exit 1
}

# Pläne laden
$plans = @()
$selectedGroups = @()

if ($UseCurrentUser) {
    # User-basierte Export: Alle Pläne des aktuellen Benutzers
    Write-PlannerLog "Verwende Pläne des aktuellen Benutzers"
    $plans = Get-AllUserPlans
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-AllUserPlans" "ERROR"
        exit 1
    }
}
elseif ($GroupIds) {
    # Direkte Angabe von Gruppen-IDs
    Write-PlannerLog "Verwende angegebene Gruppen-IDs"
    $plans = Get-PlansByGroupIds -GroupIds $GroupIds
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-PlansByGroupIds" "ERROR"
        exit 1
    }
}
elseif ($GroupNames) {
    # Suche nach Gruppennamen
    Write-PlannerLog "Suche Gruppen nach Namen"
    $selectedGroups = Get-GroupsByNames -GroupNames $GroupNames

    if ($null -eq $selectedGroups -or $selectedGroups.Count -eq 0) {
        Write-PlannerLog "Keine Gruppen mit den angegebenen Namen gefunden." "ERROR"
        exit 1
    }

    $plans = Get-PlansByGroupIds -GroupIds $selectedGroups.id
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-PlansByGroupIds" "ERROR"
        exit 1
    }
}
elseif ($Interactive) {
    # Interaktive Auswahl
    Write-PlannerLog "Starte interaktive Gruppenauswahl"
    $allGroups = Get-AllM365Groups

    if ($null -eq $allGroups -or $allGroups.Count -eq 0) {
        Write-PlannerLog "Keine M365-Gruppen verfügbar oder gefunden." "ERROR"
        exit 1
    }

    $selectedGroups = Show-GroupSelectionMenu -Groups $allGroups

    if ($null -eq $selectedGroups -or $selectedGroups.Count -eq 0) {
        Write-PlannerLog "Keine Gruppen ausgewählt. Abbruch." "WARN"
        exit 0
    }

    $plans = Get-PlansByGroupIds -GroupIds $selectedGroups.id
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-PlansByGroupIds" "ERROR"
        exit 1
    }
}
else {
    # Standardverhalten: Interaktive Auswahl anzeigen
    Write-PlannerLog "Keine Gruppen angegeben. Zeige verfügbare M365-Gruppen..." "WARN"
    Write-Host ""
    Write-Host "HINWEIS: Sie haben keine Gruppen angegeben." -ForegroundColor Yellow
    Write-Host "Verwenden Sie:" -ForegroundColor Yellow
    Write-Host "  -GroupNames 'Gruppenname'" -ForegroundColor Cyan
    Write-Host "  -GroupIds 'gruppe-id'" -ForegroundColor Cyan
    Write-Host "  -Interactive  (für interaktive Auswahl)" -ForegroundColor Cyan
    Write-Host ""

    $allGroups = Get-AllM365Groups

    if ($null -eq $allGroups -or $allGroups.Count -eq 0) {
        Write-PlannerLog "Keine M365-Gruppen verfügbar oder gefunden." "ERROR"
        exit 1
    }

    Write-Host "Möchten Sie aus den verfügbaren Gruppen auswählen? (J/N)" -ForegroundColor Yellow
    $choice = Read-Host

    if ($choice -eq "J" -or $choice -eq "j" -or $choice -eq "Y" -or $choice -eq "y") {
        $selectedGroups = Show-GroupSelectionMenu -Groups $allGroups

        if ($null -eq $selectedGroups -or $selectedGroups.Count -eq 0) {
            Write-PlannerLog "Keine Gruppen ausgewählt. Abbruch." "WARN"
            exit 0
        }

        $plans = Get-PlansByGroupIds -GroupIds $selectedGroups.id
        if ($null -eq $plans) {
            Write-PlannerLog "Fehler: Keine Rückgabe von Get-PlansByGroupIds" "ERROR"
            exit 1
        }
    }
    else {
        Write-PlannerLog "Abbruch durch Benutzer." "WARN"
        exit 0
    }
}

# Finale Validierung der Pläne
if ($null -eq $plans) {
    Write-PlannerLog "Kritischer Fehler: Plans-Variable ist null" "ERROR"
    exit 1
}

if ($plans.Count -eq 0) {
    Write-PlannerLog "Keine Pläne gefunden. Überprüfen Sie die Berechtigungen oder ob die Gruppen Pläne enthalten." "ERROR"
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
    Write-PlannerLog "=== Exportiere Plan: $($plan.title) ===" "OK"
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
try {
    $indexData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$ExportPath\_ExportIndex.json" -Encoding utf8BOM -ErrorAction Stop
}
catch {
    Write-PlannerLog "Fehler beim Schreiben der ExportIndex-Datei: $_" "ERROR"
    throw
}

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
Write-PlannerLog "Export abgeschlossen. $($plans.Count) Pläne mit $totalTasks Tasks exportiert." "OK"

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

#endregion
