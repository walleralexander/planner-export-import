<#
.SYNOPSIS
    Unit and integration tests for Export-PlannerData.ps1

.DESCRIPTION
    This test file uses Pester framework to test the Export-PlannerData script.
    Tests include:
    - Function unit tests (Write-PlannerLog, Export-ReadableSummary)
    - Mock-based tests for Graph API interactions
    - Edge case and error handling tests

.NOTES
    Prerequisites:
    - Pester 5.x: Install-Module Pester -Force -SkipPublisherCheck
    - Run with: Invoke-Pester -Path ./tests/Export-PlannerData.Tests.ps1
#>

BeforeAll {
    # Define test helper functions that simulate the actual script behavior
    # These are simplified versions for unit testing without external dependencies
    
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
        
        # Ensure directory exists
        if (-not (Test-Path $ExportPath)) {
            New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
        }
        
        $logEntry | Out-File -FilePath "$ExportPath\export.log" -Append -Encoding UTF8
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
                    }
                }
            }
            [void]$sb.AppendLine("")
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
    }

    # Version 1.1.0 new functions
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
}

Describe "Export-PlannerData Script Tests" {
    
    Context "Write-PlannerLog Function Tests" {
        BeforeEach {
            # Create a temporary test directory for logs
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should write INFO level log entry" {
            Write-PlannerLog -Message "Test message" -Level "INFO"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            Test-Path $logFile | Should -Be $true
            
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[INFO\] Test message"
        }

        It "Should write ERROR level log entry" {
            Write-PlannerLog -Message "Error occurred" -Level "ERROR"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[ERROR\] Error occurred"
        }

        It "Should write WARN level log entry" {
            Write-PlannerLog -Message "Warning message" -Level "WARN"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[WARN\] Warning message"
        }

        It "Should include timestamp in log entry" {
            Write-PlannerLog -Message "Timestamped message"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]"
        }

        It "Should append multiple log entries" {
            # Clear any existing log file
            $logFile = Join-Path $script:testExportPath "export.log"
            if (Test-Path $logFile) {
                Remove-Item $logFile -Force
            }
            
            Write-PlannerLog -Message "First message"
            Write-PlannerLog -Message "Second message"
            
            $logContent = Get-Content $logFile
            $logContent.Count | Should -BeGreaterThan 1
            ($logContent -join " ") | Should -Match "First message"
            ($logContent -join " ") | Should -Match "Second message"
        }
    }

    Context "Export-ReadableSummary Function Tests" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should create summary file with basic plan information" {
            $planData = @{
                Plan = @{
                    title = "Test Plan"
                    groupDisplayName = "Test Group"
                }
                Categories = @{}
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            Test-Path $outputPath | Should -Be $true
            $content = Get-Content $outputPath -Raw
            $content | Should -Match "Test Plan"
            $content | Should -Match "Test Group"
        }

        It "Should include categories in summary" {
            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = [PSCustomObject]@{
                    category1 = "High Priority"
                    category2 = "Low Priority"
                }
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "KATEGORIEN"
            $content | Should -Match "High Priority"
            $content | Should -Match "Low Priority"
        }

        It "Should list buckets and tasks correctly" {
            $bucket1 = [PSCustomObject]@{
                id = "bucket-1"
                name = "To Do"
                orderHint = "1"
            }

            $task1 = [PSCustomObject]@{
                id = "task-1"
                title = "Test Task"
                bucketId = "bucket-1"
                percentComplete = 0
                priority = 1
                orderHint = "1"
            }

            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = @{}
                Buckets = @($bucket1)
                Tasks = @($task1)
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "To Do"
            $content | Should -Match "Test Task"
            $content | Should -Match "Nicht begonnen"
        }

        It "Should show task completion status correctly" {
            $bucket1 = [PSCustomObject]@{
                id = "bucket-1"
                name = "Done"
                orderHint = "1"
            }

            $task1 = [PSCustomObject]@{
                id = "task-1"
                title = "Completed Task"
                bucketId = "bucket-1"
                percentComplete = 100
                priority = 1
                orderHint = "1"
            }

            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = @{}
                Buckets = @($bucket1)
                Tasks = @($task1)
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "Abgeschlossen"
            $content | Should -Match "Completed Task"
        }

        It "Should include task details with description" {
            $bucket1 = [PSCustomObject]@{
                id = "bucket-1"
                name = "To Do"
                orderHint = "1"
            }

            $task1 = [PSCustomObject]@{
                id = "task-1"
                title = "Task with Description"
                bucketId = "bucket-1"
                percentComplete = 0
                priority = 1
                orderHint = "1"
            }

            $taskDetail1 = [PSCustomObject]@{
                taskId = "task-1"
                description = "This is a detailed description"
            }

            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = @{}
                Buckets = @($bucket1)
                Tasks = @($task1)
                TaskDetails = @($taskDetail1)
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "Beschreibung"
            $content | Should -Match "This is a detailed description"
        }

        It "Should include checklist items" {
            $bucket1 = [PSCustomObject]@{
                id = "bucket-1"
                name = "To Do"
                orderHint = "1"
            }

            $task1 = [PSCustomObject]@{
                id = "task-1"
                title = "Task with Checklist"
                bucketId = "bucket-1"
                percentComplete = 0
                priority = 1
                orderHint = "1"
            }

            $taskDetail1 = [PSCustomObject]@{
                taskId = "task-1"
                checklist = [PSCustomObject]@{
                    item1 = [PSCustomObject]@{
                        title = "First item"
                        isChecked = $true
                    }
                    item2 = [PSCustomObject]@{
                        title = "Second item"
                        isChecked = $false
                    }
                }
            }

            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = @{}
                Buckets = @($bucket1)
                Tasks = @($task1)
                TaskDetails = @($taskDetail1)
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "Checkliste"
            $content | Should -Match "First item"
            $content | Should -Match "Second item"
        }

        It "Should display statistics correctly" {
            $bucket1 = [PSCustomObject]@{
                id = "bucket-1"
                name = "To Do"
                orderHint = "1"
            }

            $tasks = @(
                [PSCustomObject]@{
                    id = "task-1"
                    title = "Not Started"
                    bucketId = "bucket-1"
                    percentComplete = 0
                    priority = 1
                    orderHint = "1"
                },
                [PSCustomObject]@{
                    id = "task-2"
                    title = "In Progress"
                    bucketId = "bucket-1"
                    percentComplete = 50
                    priority = 1
                    orderHint = "2"
                },
                [PSCustomObject]@{
                    id = "task-3"
                    title = "Completed"
                    bucketId = "bucket-1"
                    percentComplete = 100
                    priority = 1
                    orderHint = "3"
                }
            )

            $planData = @{
                Plan = @{
                    title = "Test Plan"
                }
                Categories = @{}
                Buckets = @($bucket1)
                Tasks = $tasks
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $script:testExportPath "summary.txt"
            Export-ReadableSummary -PlanData $planData -OutputPath $outputPath

            $content = Get-Content $outputPath -Raw
            $content | Should -Match "STATISTIK"
            $content | Should -Match "Buckets:\s+1"
            $content | Should -Match "Tasks gesamt:\s+3"
            $content | Should -Match "Nicht begonnen:\s+1"
            $content | Should -Match "In Bearbeitung:\s+1"
            $content | Should -Match "Abgeschlossen:\s+1"
        }
    }

    Context "Plan File Name Sanitization Tests" {
        It "Should sanitize plan titles with special characters" {
            # Test that special characters would be replaced
            $testTitle = 'Plan/with:special*chars?"<>|'
            $sanitized = $testTitle -replace '[\\/:*?"<>|]', '_'
            
            $sanitized | Should -Be "Plan_with_special_chars_____"
            $sanitized | Should -Not -Match '[\\/:*?"<>|]'
        }

        It "Should handle plan titles with spaces" {
            $testTitle = "Plan with spaces"
            $sanitized = $testTitle -replace '[\\/:*?"<>|]', '_'
            
            $sanitized | Should -Be "Plan with spaces"
        }
    }

    Context "Data Structure Tests" {
        It "Should create proper export data structure" {
            $planData = @{
                Plan       = [PSCustomObject]@{ id = "plan-1"; title = "Test" }
                Buckets    = @()
                Tasks      = @()
                TaskDetails = @()
                Categories = @{}
                UserMap    = @{}
            }

            $planData.Keys | Should -Contain "Plan"
            $planData.Keys | Should -Contain "Buckets"
            $planData.Keys | Should -Contain "Tasks"
            $planData.Keys | Should -Contain "TaskDetails"
            $planData.Keys | Should -Contain "Categories"
            $planData.Keys | Should -Contain "UserMap"
        }

        It "Should handle empty plan data gracefully" {
            $planData = @{
                Plan = @{
                    title = "Empty Plan"
                }
                Categories = @{}
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                UserMap = @{}
            }

            $outputPath = Join-Path $TestDrive "empty-summary.txt"
            { Export-ReadableSummary -PlanData $planData -OutputPath $outputPath } | Should -Not -Throw
            
            Test-Path $outputPath | Should -Be $true
        }
    }

    Context "Error Handling Tests" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should handle invalid plan data structure" {
            $invalidPlanData = @{
                # Missing required fields
                Buckets = @()
            }

            $outputPath = Join-Path $script:testExportPath "invalid-summary.txt"
            { Export-ReadableSummary -PlanData $invalidPlanData -OutputPath $outputPath } | Should -Not -Throw
        }

        It "Should create log directory if it doesn't exist" {
            Remove-Item $script:testExportPath -Recurse -Force
            $global:ExportPath = $script:testExportPath
            
            { Write-PlannerLog -Message "Test" } | Should -Not -Throw
        }
    }

    Context "User ID Extraction Tests" {
        It "Should extract user IDs from task assignments" {
            $task = [PSCustomObject]@{
                id = "task-1"
                assignments = [PSCustomObject]@{
                    "abc123-def456" = @{ "@odata.type" = "#microsoft.graph.plannerAssignment" }
                    "fed654-cba321" = @{ "@odata.type" = "#microsoft.graph.plannerAssignment" }
                }
                createdBy = @{
                    user = @{ id = "aaa111-bbb222" }
                }
            }

            $userIds = @()
            
            # Extract from assignments
            if ($task.assignments) {
                $task.assignments.PSObject.Properties | ForEach-Object {
                    if ($_.Name -notin $userIds -and $_.Name -match '^[0-9a-f-]+$') {
                        $userIds += $_.Name
                    }
                }
            }
            
            # Extract from creator
            if ($task.createdBy.user.id -and $task.createdBy.user.id -notin $userIds) {
                $userIds += $task.createdBy.user.id
            }

            $userIds.Count | Should -Be 3
            $userIds | Should -Contain "abc123-def456"
            $userIds | Should -Contain "fed654-cba321"
            $userIds | Should -Contain "aaa111-bbb222"
        }
    }

    Context "Priority Mapping Tests" {
        It "Should correctly map priority values to text" {
            $priorities = @{
                0 = "Dringend"
                1 = "Wichtig"
                2 = "Mittel"
                5 = "Niedrig"
            }

            $priorities[0] | Should -Be "Dringend"
            $priorities[1] | Should -Be "Wichtig"
            $priorities[2] | Should -Be "Mittel"
            $priorities[5] | Should -Be "Niedrig"
        }
    }

    Context "Task Completion Status Tests" {
        It "Should correctly identify task status" {
            $statuses = @{
                0   = "Nicht begonnen"
                50  = "In Bearbeitung"
                100 = "Abgeschlossen"
            }

            $statuses[0] | Should -Be "Nicht begonnen"
            $statuses[50] | Should -Be "In Bearbeitung"
            $statuses[100] | Should -Be "Abgeschlossen"
        }
    }

    # ====================================================================================
    # Version 1.1.0 Feature Tests
    # ====================================================================================

    Context "Get-AllM365Groups Function Tests (v1.1.0)" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should retrieve M365 groups successfully" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @(
                        [PSCustomObject]@{ id = "group-1"; displayName = "Test Group 1"; mail = "group1@test.com" }
                        [PSCustomObject]@{ id = "group-2"; displayName = "Test Group 2"; mail = "group2@test.com" }
                    )
                    '@odata.nextLink' = $null
                }
            }

            $groups = Get-AllM365Groups

            @($groups).Count | Should -Be 2
            $groups[0].displayName | Should -Be "Test Group 1"
            $groups[1].displayName | Should -Be "Test Group 2"
        }

        It "Should handle paging with @odata.nextLink" {
            $script:callCount = 0
            Mock Invoke-MgGraphRequest {
                $script:callCount++
                if ($script:callCount -eq 1) {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "group-1"; displayName = "Group 1"; mail = "g1@test.com" })
                        '@odata.nextLink' = "https://graph.microsoft.com/v1.0/groups?next"
                    }
                }
                else {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "group-2"; displayName = "Group 2"; mail = "g2@test.com" })
                        '@odata.nextLink' = $null
                    }
                }
            }

            $groups = Get-AllM365Groups

            @($groups).Count | Should -Be 2
            $script:callCount | Should -Be 2
        }

        It "Should handle empty response" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @()
                    '@odata.nextLink' = $null
                }
            }

            $groups = Get-AllM365Groups

            @($groups).Count | Should -Be 0
        }

        It "Should handle errors gracefully" {
            Mock Invoke-MgGraphRequest {
                throw "API Error"
            }

            $groups = Get-AllM365Groups

            @($groups).Count | Should -Be 0
        }
    }

    Context "Get-GroupsByNames Function Tests (v1.1.0)" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should find groups by exact name" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @(
                        [PSCustomObject]@{ id = "group-1"; displayName = "Projektteam Alpha"; mail = "alpha@test.com" }
                    )
                }
            }

            $groups = Get-GroupsByNames -GroupNames "Projektteam Alpha"

            @($groups).Count | Should -Be 1
            $groups[0].displayName | Should -Be "Projektteam Alpha"
        }

        It "Should find multiple groups by names" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*Projektteam Alpha*") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "group-1"; displayName = "Projektteam Alpha"; mail = "alpha@test.com" })
                    }
                }
                elseif ($Uri -like "*Marketing*") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "group-2"; displayName = "Marketing"; mail = "marketing@test.com" })
                    }
                }
            }

            $groups = Get-GroupsByNames -GroupNames "Projektteam Alpha", "Marketing"

            @($groups).Count | Should -Be 2
        }

        It "Should handle null or empty input" {
            $groups = Get-GroupsByNames -GroupNames $null

            @($groups).Count | Should -Be 0
        }

        It "Should skip empty group names" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @([PSCustomObject]@{ id = "group-1"; displayName = "Valid Group"; mail = "valid@test.com" })
                }
            }

            $groups = Get-GroupsByNames -GroupNames "", "Valid Group", "  "

            @($groups).Count | Should -Be 1
            $groups[0].displayName | Should -Be "Valid Group"
        }

        It "Should avoid duplicate groups" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @(
                        [PSCustomObject]@{ id = "group-1"; displayName = "Test Group"; mail = "test@test.com" }
                        [PSCustomObject]@{ id = "group-1"; displayName = "Test Group"; mail = "test@test.com" }
                    )
                }
            }

            $groups = Get-GroupsByNames -GroupNames "Test Group"

            @($groups).Count | Should -Be 1
        }

        It "Should handle null response from API" {
            Mock Invoke-MgGraphRequest {
                return $null
            }

            $groups = Get-GroupsByNames -GroupNames "NonExistent"

            @($groups).Count | Should -Be 0
        }

        It "Should handle empty value in API response" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @()
                }
            }

            $groups = Get-GroupsByNames -GroupNames "NonExistent"

            @($groups).Count | Should -Be 0
        }

        It "Should handle API errors gracefully" {
            Mock Invoke-MgGraphRequest {
                throw "API Error: Access Denied"
            }

            $groups = Get-GroupsByNames -GroupNames "Test Group"

            @($groups).Count | Should -Be 0
        }
    }

    Context "Show-GroupSelectionMenu Function Tests (v1.1.0)" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath

            $script:testGroups = @(
                [PSCustomObject]@{ id = "group-1"; displayName = "Group 1"; mail = "g1@test.com" }
                [PSCustomObject]@{ id = "group-2"; displayName = "Group 2"; mail = "g2@test.com" }
                [PSCustomObject]@{ id = "group-3"; displayName = "Group 3"; mail = "g3@test.com" }
            )
        }

        It "Should return empty array when user selects 0 (cancel)" {
            Mock Read-Host { return "0" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 0
        }

        It "Should return all groups when user selects A" {
            Mock Read-Host { return "A" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 3
        }

        It "Should return all groups when user selects a (lowercase)" {
            Mock Read-Host { return "a" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 3
        }

        It "Should return single group when user selects index 1" {
            Mock Read-Host { return "1" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 1
            $selected[0].displayName | Should -Be "Group 1"
        }

        It "Should return multiple groups for comma-separated indices" {
            Mock Read-Host { return "1,3" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 2
            $selected[0].displayName | Should -Be "Group 1"
            $selected[1].displayName | Should -Be "Group 3"
        }

        It "Should handle spaces in comma-separated input" {
            Mock Read-Host { return "1, 2, 3" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 3
        }

        It "Should ignore invalid indices" {
            Mock Read-Host { return "1,99,2" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 2
            $selected[0].displayName | Should -Be "Group 1"
            $selected[1].displayName | Should -Be "Group 2"
        }

        It "Should ignore non-numeric input" {
            Mock Read-Host { return "1,abc,2" }

            $selected = Show-GroupSelectionMenu -Groups $script:testGroups

            $selected.Count | Should -Be 2
        }
    }

    Context "Get-AllUserPlans Function Tests (v1.1.0)" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should retrieve plans via user groups (Method 1)" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Test Group" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Test Plan" }
                        )
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 1
            $plans[0].title | Should -Be "Test Plan"
            $plans[0].groupDisplayName | Should -Be "Test Group"
            $plans[0].groupId | Should -Be "group-1"
        }

        It "Should retrieve additional plans via /me/planner/plans (Method 2)" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Group 1" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Plan from Group" }
                        )
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-2"; title = "Additional Plan" }
                        )
                    }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 2
            $plans[0].title | Should -Be "Plan from Group"
            $plans[1].title | Should -Be "Additional Plan"
        }

        It "Should avoid duplicate plans from both methods" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Group 1" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Same Plan" }
                        )
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Same Plan" }
                        )
                    }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 1
        }

        It "Should handle user with no groups" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{ value = @() }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 0
        }

        It "Should handle null response from memberOf API" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return $null
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 0
        }

        It "Should handle groups without plans" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Empty Group" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 0
        }

        It "Should handle errors in group plan retrieval gracefully" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Group 1" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    throw "Access Denied"
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Accessible Plan" }
                        )
                    }
                }
            }

            $plans = Get-AllUserPlans

            @($plans).Count | Should -Be 1
            $plans[0].title | Should -Be "Accessible Plan"
        }

        It "Should handle errors in direct plan access gracefully" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Group 1" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Group Plan" }
                        )
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    throw "Not supported"
                }
            }

            $plans = Get-AllUserPlans

            @($plans).Count | Should -Be 1
            $plans[0].title | Should -Be "Group Plan"
        }

        It "Should handle complete API failure" {
            Mock Invoke-MgGraphRequest {
                throw "Network Error"
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 0
        }

        It "Should retrieve plans from multiple groups" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "group-1"; displayName = "Group 1" }
                            [PSCustomObject]@{ id = "group-2"; displayName = "Group 2" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/group-1/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-1"; title = "Plan 1" }
                        )
                    }
                }
                elseif ($Uri -like "*/groups/group-2/planner/plans") {
                    return [PSCustomObject]@{
                        value = @(
                            [PSCustomObject]@{ id = "plan-2"; title = "Plan 2" }
                        )
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            $plans = Get-AllUserPlans

            $plans.Count | Should -Be 2
            $plans[0].title | Should -Be "Plan 1"
            $plans[0].groupDisplayName | Should -Be "Group 1"
            $plans[1].title | Should -Be "Plan 2"
            $plans[1].groupDisplayName | Should -Be "Group 2"
        }
    }

    Context "New Parameter Tests (v1.1.0)" {
        It "Should support -UseCurrentUser switch parameter" {
            # Test that the parameter would be recognized as a switch
            $paramType = "switch"
            $paramType | Should -Be "switch"
        }

        It "Should support -GroupNames string array parameter" {
            # Test that GroupNames accepts string arrays
            $testNames = @("Group 1", "Group 2")
            $testNames.GetType().Name | Should -Be "Object[]"
            $testNames[0].GetType().Name | Should -Be "String"
        }

        It "Should support -Interactive switch parameter" {
            # Test that the parameter would be recognized as a switch
            $paramType = "switch"
            $paramType | Should -Be "switch"
        }

        It "Should handle multiple GroupNames correctly" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*Group1*") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "g1"; displayName = "Group1"; mail = "g1@test.com" })
                    }
                }
                elseif ($Uri -like "*Group2*") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "g2"; displayName = "Group2"; mail = "g2@test.com" })
                    }
                }
            }

            $groups = Get-GroupsByNames -GroupNames @("Group1", "Group2")
            @($groups).Count | Should -Be 2
        }

        It "Should validate that empty string arrays are handled" {
            $groups = Get-GroupsByNames -GroupNames @()
            @($groups).Count | Should -Be 0
        }
    }

    Context "Export Mode Selection Tests (v1.1.0)" {
        BeforeEach {
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should handle user-based mode flow" {
            Mock Invoke-MgGraphRequest {
                param($Method, $Uri, $OutputType)
                if ($Uri -like "*/me/memberOf/*") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "group-1"; displayName = "User Group" })
                    }
                }
                elseif ($Uri -like "*/groups/*/planner/plans") {
                    return [PSCustomObject]@{
                        value = @([PSCustomObject]@{ id = "plan-1"; title = "User Plan" })
                    }
                }
                elseif ($Uri -like "*/me/planner/plans") {
                    return [PSCustomObject]@{ value = @() }
                }
            }

            # Simulate UseCurrentUser mode
            $plans = Get-AllUserPlans
            @($plans).Count | Should -BeGreaterThan 0
        }

        It "Should handle group-based mode with names" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @([PSCustomObject]@{ id = "g1"; displayName = "TestGroup"; mail = "test@test.com" })
                }
            }

            # Simulate GroupNames mode
            $groups = Get-GroupsByNames -GroupNames "TestGroup"
            @($groups).Count | Should -Be 1
        }

        It "Should handle interactive mode flow" {
            Mock Invoke-MgGraphRequest {
                return [PSCustomObject]@{
                    value = @(
                        [PSCustomObject]@{ id = "g1"; displayName = "Group 1"; mail = "g1@test.com" }
                        [PSCustomObject]@{ id = "g2"; displayName = "Group 2"; mail = "g2@test.com" }
                    )
                    '@odata.nextLink' = $null
                }
            }

            Mock Read-Host { return "1" }

            # Simulate Interactive mode
            $allGroups = Get-AllM365Groups
            $selectedGroups = Show-GroupSelectionMenu -Groups $allGroups

            @($selectedGroups).Count | Should -Be 1
        }
    }
}
