
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath = "C:\planner-data\PlannerExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupIds,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupNames,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCompletedTasks,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$UseCurrentUser
)

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
        $logEntry | Out-File -FilePath "$ExportPath\export.log" -Append -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-Host "[ERROR] Konnte nicht in Log-Datei schreiben: $_" -ForegroundColor Red
    }
}

function Test-SafePath {
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

if ([string]::IsNullOrWhiteSpace($Path)) {
        if ($ErrorMessage) { $ErrorMessage.Value = "Pfad darf nicht leer sein" }
        return $false
    }

if ($Path -match '^\\\\') {
        if ($ErrorMessage) { $ErrorMessage.Value = "UNC-Pfade (Netzwerkpfade) sind aus Sicherheitsgründen nicht erlaubt: $Path" }
        return $false
    }

try {
        $normalizedPath = [System.IO.Path]::GetFullPath($Path)
    }
    catch {
        if ($ErrorMessage) { $ErrorMessage.Value = "Ungültiges Pfad-Format: $($_.Exception.Message)" }
        return $false
    }

if ($Mode -eq 'Export') {
        
        if (Test-Path $normalizedPath) {
            
            if (-not (Test-Path $normalizedPath -PathType Container)) {
                if ($ErrorMessage) { $ErrorMessage.Value = "Pfad existiert bereits als Datei (kein Verzeichnis): $normalizedPath" }
                return $false
            }

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
            
            $parentPath = Split-Path $normalizedPath -Parent

            if (-not $parentPath) {
                if ($ErrorMessage) { $ErrorMessage.Value = "Kann übergeordnetes Verzeichnis nicht ermitteln für: $normalizedPath" }
                return $false
            }

            if (-not (Test-Path $parentPath)) {
                if ($AllowCreate) {
                    
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
        
        if (-not (Test-Path $normalizedPath)) {
            if ($ErrorMessage) { $ErrorMessage.Value = "Import-Verzeichnis existiert nicht: $normalizedPath" }
            return $false
        }

if (-not (Test-Path $normalizedPath -PathType Container)) {
            if ($ErrorMessage) { $ErrorMessage.Value = "Import-Pfad ist kein Verzeichnis: $normalizedPath" }
            return $false
        }

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

function Connect-ToGraph {
    Write-PlannerLog "Verbinde mit Microsoft Graph..."
    try {
        
        $context = Get-MgContext
        if ($null -eq $context) {
            try {
                
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -NoWelcome -ErrorAction Stop
            }
            catch {
                
                Write-PlannerLog "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." "WARN"
                Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "Tasks.ReadWrite", "User.Read", "User.ReadBasic.All" -UseDeviceCode -NoWelcome
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

$uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(g:g eq 'Unified') and (displayName eq '$name' or startswith(displayName,'$name'))&`$select=id,displayName,mail"

            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

            if ($null -eq $response) {
                Write-PlannerLog "  Keine Antwort vom Server für Gruppe: $name" "WARN"
            }
            elseif ($response.value -and $response.value.Count -gt 0) {
                $matchCount = $response.value.Count
                Write-PlannerLog "  $matchCount Gruppe(n) gefunden für: $name"
                foreach ($group in $response.value) {
                    
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

try {
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/details" -OutputType PSObject
        $planData.Categories = $planDetails.categoryDescriptions
        $planData["PlanDetails"] = $planDetails

if ($planDetails.'@odata.etag') {
            $planData["PlanDetailsEtag"] = $planDetails.'@odata.etag'
        }
    }
    catch {
        Write-PlannerLog "  Konnte Plan-Details nicht laden: $_" "WARN"
    }

try {
        $buckets = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/plans/$($Plan.id)/buckets" -OutputType PSObject
        $planData.Buckets = $buckets.value
        Write-PlannerLog "  $($buckets.value.Count) Buckets geladen"
    }
    catch {
        Write-PlannerLog "  Fehler beim Laden der Buckets: $_" "ERROR"
    }

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

if (-not $IncludeCompletedTasks) {
            $completedCount = ($allTasks | Where-Object { $_.percentComplete -eq 100 }).Count
            Write-PlannerLog "  $completedCount abgeschlossene Tasks werden übersprungen (verwende -IncludeCompletedTasks um diese einzubeziehen)"
            $allTasks = $allTasks | Where-Object { $_.percentComplete -ne 100 }
        }

        $planData.Tasks = $allTasks
        Write-PlannerLog "  $($allTasks.Count) Tasks geladen"

$taskDetails = @()
        $counter = 0
        foreach ($task in $allTasks) {
            $counter++
            Write-Progress -Activity "Lade Task-Details für '$($Plan.title)'" -Status "Task $counter von $($allTasks.Count)" -PercentComplete (($counter / $allTasks.Count) * 100)

            try {
                $detail = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details" -OutputType PSObject
                $detail | Add-Member -NotePropertyName "taskId" -NotePropertyValue $task.id -Force
                $taskDetails += $detail

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
            Write-PlannerLog "    Konnte Benutzer $userId nicht auflösen" "WARN"
            $userMap[$userId] = @{ Id = $userId; DisplayName = "Unbekannt"; UserPrincipalName = ""; Mail = "" }
        }
    }
    $planData["UserMap"] = $userMap

$sanitizedTitle = ($Plan.title -replace '[\\/:*?"<>|]', '_')
    
    if ($sanitizedTitle.Length -gt 100) {
        $sanitizedTitle = $sanitizedTitle.Substring(0, 100)
    }
    
    $planFileName = $sanitizedTitle.TrimEnd(' ', '.')
    
    if ([string]::IsNullOrWhiteSpace($planFileName)) {
        $planFileName = "unnamed_plan_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }
    $planFilePath = Join-Path $PlanExportPath "$planFileName.json"

    try {
        $planData | ConvertTo-Json -Depth 20 | Out-File -FilePath $planFilePath -Encoding utf8 -ErrorAction Stop
    }
    catch {
        Write-PlannerLog "Fehler beim Schreiben der JSON-Datei: $planFilePath - $_" "ERROR"
        throw
    }
    Write-PlannerLog "  Plan exportiert nach: $planFilePath" "OK"

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

if ($PlanData.Categories) {
        [void]$sb.AppendLine("KATEGORIEN / LABELS:")
        $PlanData.Categories.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object {
            [void]$sb.AppendLine("  $($_.Name): $($_.Value)")
        }
        [void]$sb.AppendLine("")
    }

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

$priority = switch ($task.priority) {
                    0 { "Dringend" }
                    1 { "Wichtig" }
                    2 { "Mittel" }
                    default { "Niedrig" }
                }
                [void]$sb.AppendLine("    Priorität: $priority")

if ($task.dueDateTime) {
                    $dueDate = [DateTime]::Parse($task.dueDateTime).ToString("dd.MM.yyyy")
                    [void]$sb.AppendLine("    Fällig am: $dueDate")
                }

if ($task.startDateTime) {
                    $startDate = [DateTime]::Parse($task.startDateTime).ToString("dd.MM.yyyy")
                    [void]$sb.AppendLine("    Start: $startDate")
                }

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

$detail = $PlanData.TaskDetails | Where-Object { $_.taskId -eq $task.id }
                if ($detail) {
                    
                    if ($detail.description) {
                        [void]$sb.AppendLine("    Beschreibung:")
                        $detail.description -split "`n" | ForEach-Object {
                            [void]$sb.AppendLine("      $_")
                        }
                    }

if ($detail.checklist) {
                        [void]$sb.AppendLine("    Checkliste:")
                        $detail.checklist.PSObject.Properties | ForEach-Object {
                            $checkItem = $_.Value
                            $checkMark = if ($checkItem.isChecked) { "[x]" } else { "[ ]" }
                            [void]$sb.AppendLine("      $checkMark $($checkItem.title)")
                        }
                    }

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

$orphanTasks = $PlanData.Tasks | Where-Object { -not $_.bucketId -or ($PlanData.Buckets.id -notcontains $_.bucketId) }
    if ($orphanTasks.Count -gt 0) {
        [void]$sb.AppendLine("-" * 60)
        [void]$sb.AppendLine("TASKS OHNE BUCKET:")
        [void]$sb.AppendLine("-" * 60)
        foreach ($task in $orphanTasks) {
            [void]$sb.AppendLine("  - $($task.title)")
        }
    }

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
        $sb.ToString() | Out-File -FilePath $OutputPath -Encoding utf8 -ErrorAction Stop
        Write-PlannerLog "  Zusammenfassung erstellt: $OutputPath" "OK"
    }
    catch {
        Write-PlannerLog "Fehler beim Schreiben der Zusammenfassung: $OutputPath - $_" "ERROR"
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Planner Export Tool" -ForegroundColor Cyan
Write-Host "  by Alexander Waller" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$pathValidationError = $null
$isDefaultPattern = $ExportPath -match '^[A-Z]:\\planner-data\\PlannerExport_\d{8}_\d{6}$'

if ($isDefaultPattern) {
    
    if (-not (Test-SafePath -Path $ExportPath -Mode Export -AllowCreate -ErrorMessage ([ref]$pathValidationError))) {
        Write-Host ""
        Write-Host "Fehler: $pathValidationError" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
}
else {
    
    if (-not (Test-SafePath -Path $ExportPath -Mode Export -ErrorMessage ([ref]$pathValidationError))) {
        Write-Host ""
        Write-Host "Fehler: $pathValidationError" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
}

if (-not (Test-Path $ExportPath)) {
    try {
        New-Item -ItemType Directory -Path $ExportPath -ErrorAction Stop | Out-Null
        Write-PlannerLog "Export-Verzeichnis erstellt: $ExportPath"
    }
    catch {
        Write-PlannerLog "Fehler beim Erstellen des Export-Verzeichnisses: $_" "ERROR"
        exit 1
    }
}
else {
    Write-PlannerLog "Verwende existierendes Export-Verzeichnis: $ExportPath"
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Planner)) {
    Write-PlannerLog "Microsoft.Graph Module werden installiert..." "WARN"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue

if (-not (Connect-ToGraph)) {
    Write-PlannerLog "Abbruch: Keine Verbindung zu Microsoft Graph möglich." "ERROR"
    exit 1
}

$plans = @()
$selectedGroups = @()

if ($UseCurrentUser) {
    
    Write-PlannerLog "Verwende Pläne des aktuellen Benutzers"
    $plans = Get-AllUserPlans
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-AllUserPlans" "ERROR"
        exit 1
    }
}
elseif ($GroupIds) {
    
    Write-PlannerLog "Verwende angegebene Gruppen-IDs"
    $plans = Get-PlansByGroupIds -GroupIds $GroupIds
    if ($null -eq $plans) {
        Write-PlannerLog "Fehler: Keine Rückgabe von Get-PlansByGroupIds" "ERROR"
        exit 1
    }
}
elseif ($GroupNames) {
    
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

if ($null -eq $plans) {
    Write-PlannerLog "Kritischer Fehler: Plans-Variable ist null" "ERROR"
    exit 1
}

if ($plans.Count -eq 0) {
    Write-PlannerLog "Keine Pläne gefunden. Überprüfen Sie die Berechtigungen oder ob die Gruppen Pläne enthalten." "ERROR"
    exit 1
}

Write-Host ""
Write-Host "Gefundene Pläne:" -ForegroundColor Yellow
for ($i = 0; $i -lt $plans.Count; $i++) {
    $groupName = if ($plans[$i].groupDisplayName) { " (Gruppe: $($plans[$i].groupDisplayName))" } else { "" }
    Write-Host "  [$($i+1)] $($plans[$i].title)$groupName"
}
Write-Host ""

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

$indexData = @{
    ExportDate    = (Get-Date).ToString("o")
    ExportedBy    = (Get-MgContext).Account
    TotalPlans    = $plans.Count
    Plans         = $exportSummary
    ScriptVersion = "1.0.0"
}
try {
    $indexData | ConvertTo-Json -Depth 10 | Out-File -FilePath "$ExportPath\_ExportIndex.json" -Encoding utf8 -ErrorAction Stop
}
catch {
    Write-PlannerLog "Fehler beim Schreiben der ExportIndex-Datei: $_" "ERROR"
    throw
}

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

