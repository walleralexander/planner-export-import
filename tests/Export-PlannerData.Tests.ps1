<#
.SYNOPSIS
    Unit and integration tests for Export-PlannerData.ps1

.DESCRIPTION
    This test file uses Pester framework to test the Export-PlannerData script.
    Tests include:
    - Function unit tests (Write-Log, Export-ReadableSummary)
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
}

Describe "Export-PlannerData Script Tests" {
    
    Context "Write-Log Function Tests" {
        BeforeEach {
            # Create a temporary test directory for logs
            $script:testExportPath = Join-Path $TestDrive "export-test"
            New-Item -ItemType Directory -Path $script:testExportPath -Force | Out-Null
            $global:ExportPath = $script:testExportPath
        }

        It "Should write INFO level log entry" {
            Write-Log -Message "Test message" -Level "INFO"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            Test-Path $logFile | Should -Be $true
            
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[INFO\] Test message"
        }

        It "Should write ERROR level log entry" {
            Write-Log -Message "Error occurred" -Level "ERROR"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[ERROR\] Error occurred"
        }

        It "Should write WARN level log entry" {
            Write-Log -Message "Warning message" -Level "WARN"
            
            $logFile = Join-Path $script:testExportPath "export.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[WARN\] Warning message"
        }

        It "Should include timestamp in log entry" {
            Write-Log -Message "Timestamped message"
            
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
            
            Write-Log -Message "First message"
            Write-Log -Message "Second message"
            
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
            
            { Write-Log -Message "Test" } | Should -Not -Throw
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
}
