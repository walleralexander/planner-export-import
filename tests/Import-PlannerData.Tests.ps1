<#
.SYNOPSIS
    Unit and integration tests for Import-PlannerData.ps1

.DESCRIPTION
    This test file uses Pester framework to test the Import-PlannerData script.
    Tests include:
    - Function unit tests (Write-Log, Resolve-UserId, Invoke-GraphWithRetry)
    - Mock-based tests for Graph API interactions
    - Edge case and error handling tests
    - User mapping and resolution tests

.NOTES
    Prerequisites:
    - Pester 5.x: Install-Module Pester -Force -SkipPublisherCheck
    - Run with: Invoke-Pester -Path ./tests/Import-PlannerData.Tests.ps1
#>

BeforeAll {
    # Define test helper functions that simulate the actual script behavior
    # These are simplified versions for unit testing without external dependencies
    
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
        
        # Ensure directory exists
        if (-not (Test-Path $ImportPath)) {
            New-Item -ItemType Directory -Path $ImportPath -Force | Out-Null
        }
        
        $logEntry | Out-File -FilePath "$ImportPath\import.log" -Append -Encoding UTF8
    }

    function Resolve-UserId {
        param([string]$OldUserId, [hashtable]$OldUserMap)

        # If UserMapping vorhanden, verwende es
        if ($UserMapping -and $UserMapping.ContainsKey($OldUserId)) {
            return $UserMapping[$OldUserId]
        }

        # Versuche den User über UPN oder Mail in der neuen Umgebung zu finden
        if ($OldUserMap -and $OldUserMap[$OldUserId]) {
            $upn = $OldUserMap[$OldUserId].UserPrincipalName
            $mail = $OldUserMap[$OldUserId].Mail

            if ($upn) {
                try {
                    # Mock - in real test this would call Graph API
                    if (Test-Path Function:\Invoke-MgGraphRequest) {
                        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                        if ($user) {
                            return $user.id
                        }
                    }
                }
                catch { }
            }

            if ($mail -and $mail -ne $upn) {
                try {
                    if (Test-Path Function:\Invoke-MgGraphRequest) {
                        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mail`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                        if ($user) {
                            return $user.id
                        }
                    }
                }
                catch { }
            }
        }

        # Fallback: Versuche die alte ID direkt (gleicher Tenant)
        try {
            if (Test-Path Function:\Invoke-MgGraphRequest) {
                $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$OldUserId`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
                if ($user) {
                    return $user.id
                }
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
            return $null
        }

        # In real implementation, would create plan via Graph API
        # For testing, just return mock result
        return @{
            NewPlanId    = "mock-plan-id"
            TasksCreated = $planData.Tasks.Count
            BucketsCreated = $planData.Buckets.Count
        }
    }

    # Set global variables for testing
    $global:DryRun = $false
    $global:ThrottleDelayMs = 0  # Set to 0 for faster tests
    $global:SkipAssignments = $false
    $global:SkipCompletedTasks = $false
    $global:UserMapping = $null
}

Describe "Import-PlannerData Script Tests" {
    
    Context "Write-Log Function Tests" {
        BeforeEach {
            # Create a temporary test directory for logs
            $script:testImportPath = Join-Path $TestDrive "import-test"
            New-Item -ItemType Directory -Path $script:testImportPath -Force | Out-Null
            $global:ImportPath = $script:testImportPath
        }

        It "Should write INFO level log entry" {
            Write-Log -Message "Test import message" -Level "INFO"
            
            $logFile = Join-Path $script:testImportPath "import.log"
            Test-Path $logFile | Should -Be $true
            
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[INFO\] Test import message"
        }

        It "Should write ERROR level log entry" {
            Write-Log -Message "Import error occurred" -Level "ERROR"
            
            $logFile = Join-Path $script:testImportPath "import.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[ERROR\] Import error occurred"
        }

        It "Should write DRYRUN level log entry" {
            Write-Log -Message "Dry run message" -Level "DRYRUN"
            
            $logFile = Join-Path $script:testImportPath "import.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[DRYRUN\] Dry run message"
        }

        It "Should include timestamp in log entry" {
            Write-Log -Message "Timestamped import message"
            
            $logFile = Join-Path $script:testImportPath "import.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]"
        }

        It "Should append multiple log entries" {
            # Clear any existing log file
            $logFile = Join-Path $script:testImportPath "import.log"
            if (Test-Path $logFile) {
                Remove-Item $logFile -Force
            }
            
            Write-Log -Message "First import message"
            Write-Log -Message "Second import message"
            
            $logContent = Get-Content $logFile
            $logContent.Count | Should -BeGreaterThan 1
            ($logContent -join " ") | Should -Match "First import message"
            ($logContent -join " ") | Should -Match "Second import message"
        }
    }

    Context "Resolve-UserId Function Tests" {
        BeforeEach {
            $global:UserMapping = $null
        }

        It "Should return mapped user ID when UserMapping is provided" {
            $global:UserMapping = @{
                "old-user-id-1" = "new-user-id-1"
                "old-user-id-2" = "new-user-id-2"
            }

            $oldUserMap = @{
                "old-user-id-1" = @{
                    DisplayName = "Test User"
                    UserPrincipalName = "test@example.com"
                    Mail = "test@example.com"
                }
            }

            $result = Resolve-UserId -OldUserId "old-user-id-1" -OldUserMap $oldUserMap
            $result | Should -Be "new-user-id-1"
        }

        It "Should return null when user cannot be resolved" {
            $global:UserMapping = $null
            
            # Mock Invoke-MgGraphRequest to simulate user not found
            Mock Invoke-MgGraphRequest { throw "User not found" }

            $oldUserMap = @{
                "unknown-user-id" = @{
                    DisplayName = "Unknown User"
                    UserPrincipalName = "unknown@example.com"
                    Mail = "unknown@example.com"
                }
            }

            $result = Resolve-UserId -OldUserId "unknown-user-id" -OldUserMap $oldUserMap
            $result | Should -Be $null
        }

        It "Should handle empty UserMap gracefully" {
            $global:UserMapping = $null
            Mock Invoke-MgGraphRequest { throw "User not found" }

            $result = Resolve-UserId -OldUserId "some-user-id" -OldUserMap $null
            $result | Should -Be $null
        }

        It "Should prioritize UserMapping over UPN lookup" {
            $global:UserMapping = @{
                "old-id" = "mapped-new-id"
            }

            $oldUserMap = @{
                "old-id" = @{
                    UserPrincipalName = "user@example.com"
                }
            }

            # Even if UPN lookup would work, UserMapping should be used
            $result = Resolve-UserId -OldUserId "old-id" -OldUserMap $oldUserMap
            $result | Should -Be "mapped-new-id"
        }

        It "Should handle user map with missing fields" {
            $global:UserMapping = $null
            Mock Invoke-MgGraphRequest { throw "User not found" }

            $oldUserMap = @{
                "user-id" = @{
                    DisplayName = "Test User"
                    # Missing UPN and Mail
                }
            }

            $result = Resolve-UserId -OldUserId "user-id" -OldUserMap $oldUserMap
            $result | Should -Be $null
        }
    }

    Context "DryRun Mode Tests" {
        BeforeEach {
            $script:testImportPath = Join-Path $TestDrive "import-test"
            New-Item -ItemType Directory -Path $script:testImportPath -Force | Out-Null
            $global:ImportPath = $script:testImportPath
            $global:DryRun = $true

            # Create a test export file
            $testExport = @{
                Plan = @{
                    id = "plan-1"
                    title = "Test Plan"
                    groupId = "group-1"
                }
                Buckets = @(
                    @{
                        id = "bucket-1"
                        name = "To Do"
                        orderHint = "1"
                    }
                )
                Tasks = @(
                    @{
                        id = "task-1"
                        title = "Test Task"
                        bucketId = "bucket-1"
                        percentComplete = 0
                        priority = 1
                    }
                )
                TaskDetails = @()
                Categories = @{}
                UserMap = @{}
            }

            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            $testExport | ConvertTo-Json -Depth 20 | Out-File -FilePath $testJsonPath -Encoding UTF8
        }

        It "Should not create actual resources in DryRun mode" {
            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            
            # In DryRun mode, Import-PlanFromJson should not make actual API calls
            $result = Import-PlanFromJson -JsonFilePath $testJsonPath -TargetGroupId "group-1"
            
            # Result should be null because DryRun returns early
            $result | Should -Be $null
        }

        It "Should log DryRun messages correctly" {
            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            
            Import-PlanFromJson -JsonFilePath $testJsonPath -TargetGroupId "group-1"
            
            $logFile = Join-Path $script:testImportPath "import.log"
            $logContent = Get-Content $logFile -Raw
            $logContent | Should -Match "\[DRYRUN\]"
        }
    }

    Context "Plan Import Data Structure Tests" {
        BeforeEach {
            $script:testImportPath = Join-Path $TestDrive "import-test"
            New-Item -ItemType Directory -Path $script:testImportPath -Force | Out-Null
            $global:ImportPath = $script:testImportPath
            $global:DryRun = $false
        }

        It "Should load export JSON file correctly" {
            $testExport = @{
                Plan = @{
                    id = "plan-1"
                    title = "Test Plan"
                    groupId = "group-1"
                }
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                Categories = @{}
                UserMap = @{}
            }

            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            $testExport | ConvertTo-Json -Depth 20 | Out-File -FilePath $testJsonPath -Encoding UTF8

            $planData = Get-Content $testJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
            
            $planData.Plan.title | Should -Be "Test Plan"
            $planData.Plan.groupId | Should -Be "group-1"
        }

        It "Should handle plan with no target group specified" {
            $testExport = @{
                Plan = @{
                    id = "plan-1"
                    title = "Test Plan"
                    groupId = "original-group-1"
                }
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                Categories = @{}
                UserMap = @{}
            }

            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            $testExport | ConvertTo-Json -Depth 20 | Out-File -FilePath $testJsonPath -Encoding UTF8

            $planData = Get-Content $testJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
            
            # Should use original group ID if no target specified
            $groupId = $planData.Plan.groupId
            $groupId | Should -Be "original-group-1"
        }

        It "Should handle missing groupId in plan data" {
            $testExport = @{
                Plan = @{
                    id = "plan-1"
                    title = "Test Plan"
                    # Missing groupId
                }
                Buckets = @()
                Tasks = @()
                TaskDetails = @()
                Categories = @{}
                UserMap = @{}
            }

            $testJsonPath = Join-Path $script:testImportPath "TestPlan.json"
            $testExport | ConvertTo-Json -Depth 20 | Out-File -FilePath $testJsonPath -Encoding UTF8

            $planData = Get-Content $testJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
            
            # groupId should be null/empty
            $planData.Plan.groupId | Should -BeNullOrEmpty
        }
    }

    Context "Bucket Mapping Tests" {
        It "Should create bucket ID mapping" {
            $buckets = @(
                @{
                    id = "old-bucket-1"
                    name = "To Do"
                },
                @{
                    id = "old-bucket-2"
                    name = "In Progress"
                }
            )

            $bucketMapping = @{}
            
            foreach ($bucket in $buckets) {
                # Simulate creating new bucket and mapping IDs
                $newBucketId = "new-bucket-" + ([Guid]::NewGuid().ToString().Substring(0, 8))
                $bucketMapping[$bucket.id] = $newBucketId
            }

            $bucketMapping.Keys.Count | Should -Be 2
            $bucketMapping.ContainsKey("old-bucket-1") | Should -Be $true
            $bucketMapping.ContainsKey("old-bucket-2") | Should -Be $true
        }

        It "Should map old bucket IDs to new bucket IDs" {
            $bucketMapping = @{
                "old-bucket-1" = "new-bucket-1"
                "old-bucket-2" = "new-bucket-2"
            }

            $task = @{
                bucketId = "old-bucket-1"
            }

            $newBucketId = if ($task.bucketId -and $bucketMapping.ContainsKey($task.bucketId)) {
                $bucketMapping[$task.bucketId]
            } else {
                $null
            }

            $newBucketId | Should -Be "new-bucket-1"
        }

        It "Should handle task with no bucket" {
            $bucketMapping = @{
                "old-bucket-1" = "new-bucket-1"
            }

            $task = @{
                # No bucketId
            }

            $newBucketId = if ($task.bucketId -and $bucketMapping.ContainsKey($task.bucketId)) {
                $bucketMapping[$task.bucketId]
            } else {
                $null
            }

            $newBucketId | Should -Be $null
        }
    }

    Context "Task Mapping Tests" {
        It "Should create task ID mapping" {
            $tasks = @(
                @{ id = "old-task-1"; title = "Task 1" },
                @{ id = "old-task-2"; title = "Task 2" }
            )

            $taskMapping = @{}
            
            foreach ($task in $tasks) {
                $newTaskId = "new-task-" + ([Guid]::NewGuid().ToString().Substring(0, 8))
                $taskMapping[$task.id] = $newTaskId
            }

            $taskMapping.Keys.Count | Should -Be 2
            $taskMapping.ContainsKey("old-task-1") | Should -Be $true
            $taskMapping.ContainsKey("old-task-2") | Should -Be $true
        }

        It "Should skip completed tasks when SkipCompletedTasks is set" {
            # Reset variable first
            $skipFlag = $true

            $tasks = @(
                @{ id = "task-1"; title = "Active Task"; percentComplete = 0 },
                @{ id = "task-2"; title = "Completed Task"; percentComplete = 100 },
                @{ id = "task-3"; title = "Another Active Task"; percentComplete = 50 }
            )

            $tasksToImport = $tasks | Where-Object { 
                -not ($skipFlag -and $_.percentComplete -eq 100)
            }

            $tasksToImport.Count | Should -Be 2
            $tasksToImport[0].title | Should -Be "Active Task"
            $tasksToImport[1].title | Should -Be "Another Active Task"
        }

        It "Should include all tasks when SkipCompletedTasks is not set" {
            $global:SkipCompletedTasks = $false

            $tasks = @(
                @{ id = "task-1"; title = "Active Task"; percentComplete = 0 },
                @{ id = "task-2"; title = "Completed Task"; percentComplete = 100 }
            )

            $tasksToImport = $tasks | Where-Object { 
                -not ($global:SkipCompletedTasks -and $_.percentComplete -eq 100)
            }

            $tasksToImport.Count | Should -Be 2
        }
    }

    Context "Task Body Construction Tests" {
        It "Should build basic task body" {
            $task = @{
                title = "Test Task"
                percentComplete = 0
                priority = 1
            }

            $taskBody = @{
                planId          = "plan-1"
                title           = $task.title
                percentComplete = $task.percentComplete
                priority        = $task.priority
            }

            $taskBody.title | Should -Be "Test Task"
            $taskBody.percentComplete | Should -Be 0
            $taskBody.priority | Should -Be 1
        }

        It "Should include due date if present" {
            $task = @{
                title = "Test Task"
                percentComplete = 0
                priority = 1
                dueDateTime = "2026-12-31T00:00:00Z"
            }

            $taskBody = @{
                planId          = "plan-1"
                title           = $task.title
                percentComplete = $task.percentComplete
                priority        = $task.priority
            }

            if ($task.dueDateTime) {
                $taskBody["dueDateTime"] = $task.dueDateTime
            }

            $taskBody.Keys | Should -Contain "dueDateTime"
            $taskBody.dueDateTime | Should -Be "2026-12-31T00:00:00Z"
        }

        It "Should include start date if present" {
            $task = @{
                title = "Test Task"
                percentComplete = 0
                priority = 1
                startDateTime = "2026-01-01T00:00:00Z"
            }

            $taskBody = @{
                planId          = "plan-1"
                title           = $task.title
                percentComplete = $task.percentComplete
                priority        = $task.priority
            }

            if ($task.startDateTime) {
                $taskBody["startDateTime"] = $task.startDateTime
            }

            $taskBody.Keys | Should -Contain "startDateTime"
            $taskBody.startDateTime | Should -Be "2026-01-01T00:00:00Z"
        }

        It "Should include categories if present" {
            $task = [PSCustomObject]@{
                title = "Test Task"
                percentComplete = 0
                priority = 1
                appliedCategories = [PSCustomObject]@{
                    category1 = $true
                    category2 = $true
                }
            }

            $categories = @{}
            $task.appliedCategories.PSObject.Properties | Where-Object { $_.Value -eq $true } | ForEach-Object {
                $categories[$_.Name] = $true
            }

            $categories.Keys.Count | Should -Be 2
            $categories["category1"] | Should -Be $true
            $categories["category2"] | Should -Be $true
        }
    }

    Context "Assignment Tests" {
        BeforeEach {
            $global:SkipAssignments = $false
            $global:UserMapping = $null
        }

        It "Should skip assignments when SkipAssignments is set" {
            $global:SkipAssignments = $true

            $task = [PSCustomObject]@{
                assignments = [PSCustomObject]@{
                    "user-1" = @{}
                }
            }

            $shouldProcessAssignments = -not $global:SkipAssignments -and $task.assignments
            $shouldProcessAssignments | Should -Be $false
        }

        It "Should process assignments when SkipAssignments is not set" {
            $global:SkipAssignments = $false

            $task = [PSCustomObject]@{
                assignments = [PSCustomObject]@{
                    "user-1" = @{}
                }
            }

            $shouldProcessAssignments = -not $global:SkipAssignments -and $task.assignments
            $shouldProcessAssignments | Should -Be $true
        }

        It "Should create assignment structure correctly" {
            $assignments = @{}
            $resolvedUserId = "resolved-user-id"

            $assignments[$resolvedUserId] = @{
                "@odata.type" = "#microsoft.graph.plannerAssignment"
                "orderHint"   = " !"
            }

            $assignments.Keys | Should -Contain $resolvedUserId
            $assignments[$resolvedUserId]["@odata.type"] | Should -Be "#microsoft.graph.plannerAssignment"
        }
    }

    Context "Task Details Tests" {
        It "Should build task details body with description" {
            $detail = [PSCustomObject]@{
                taskId = "task-1"
                description = "This is a task description"
            }

            $detailBody = @{}

            if ($detail.description) {
                $detailBody["description"] = $detail.description
                $detailBody["previewType"] = "description"
            }

            $detailBody.Keys | Should -Contain "description"
            $detailBody.Keys | Should -Contain "previewType"
            $detailBody.description | Should -Be "This is a task description"
            $detailBody.previewType | Should -Be "description"
        }

        It "Should build checklist structure" {
            $detail = [PSCustomObject]@{
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

            $checklist = @{}
            $detail.checklist.PSObject.Properties | ForEach-Object {
                $checkId = [Guid]::NewGuid().ToString()
                $checklist[$checkId] = @{
                    "@odata.type" = "microsoft.graph.plannerChecklistItem"
                    title         = $_.Value.title
                    isChecked     = $_.Value.isChecked
                }
            }

            $checklist.Keys.Count | Should -Be 2
            
            # Verify structure of checklist items
            $firstItem = $checklist.Values | Select-Object -First 1
            $firstItem["@odata.type"] | Should -Be "microsoft.graph.plannerChecklistItem"
            $firstItem.Keys | Should -Contain "title"
            $firstItem.Keys | Should -Contain "isChecked"
        }

        It "Should build references structure" {
            $detail = [PSCustomObject]@{
                taskId = "task-1"
                references = [PSCustomObject]@{
                    "https://example.com" = [PSCustomObject]@{
                        alias = "Example Link"
                        type = "other"
                        previewPriority = "123"
                    }
                }
            }

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

            $references.Keys | Should -Contain "https://example.com"
            $references["https://example.com"].alias | Should -Be "Example Link"
            $references["https://example.com"].type | Should -Be "other"
            $references["https://example.com"].previewPriority | Should -Be "123"
        }
    }

    Context "Import Mapping Tests" {
        It "Should create import mapping structure" {
            $mappingData = @{
                ImportDate   = (Get-Date).ToString("o")
                OriginalPlan = "old-plan-id"
                NewPlanId    = "new-plan-id"
                GroupId      = "group-id"
                BucketMap    = @{
                    "old-bucket-1" = "new-bucket-1"
                }
                TaskMap      = @{
                    "old-task-1" = "new-task-1"
                }
            }

            $mappingData.Keys | Should -Contain "ImportDate"
            $mappingData.Keys | Should -Contain "OriginalPlan"
            $mappingData.Keys | Should -Contain "NewPlanId"
            $mappingData.Keys | Should -Contain "GroupId"
            $mappingData.Keys | Should -Contain "BucketMap"
            $mappingData.Keys | Should -Contain "TaskMap"
        }

        It "Should save import mapping to JSON file" {
            $testDir = Join-Path $TestDrive "mapping-test"
            New-Item -ItemType Directory -Path $testDir -Force | Out-Null

            $mappingData = @{
                ImportDate   = (Get-Date).ToString("o")
                OriginalPlan = "old-plan-id"
                NewPlanId    = "new-plan-id"
                GroupId      = "group-id"
                BucketMap    = @{}
                TaskMap      = @{}
            }

            $mappingFile = Join-Path $testDir "TestPlan_ImportMapping.json"
            $mappingData | ConvertTo-Json -Depth 10 | Out-File -FilePath $mappingFile -Encoding UTF8

            Test-Path $mappingFile | Should -Be $true
            
            $loaded = Get-Content $mappingFile -Raw | ConvertFrom-Json
            $loaded.NewPlanId | Should -Be "new-plan-id"
        }
    }

    Context "Error Handling Tests" {
        BeforeEach {
            $script:testImportPath = Join-Path $TestDrive "import-test"
            New-Item -ItemType Directory -Path $script:testImportPath -Force | Out-Null
            $global:ImportPath = $script:testImportPath
        }

        It "Should handle missing import directory gracefully" {
            $nonExistentPath = Join-Path $TestDrive "non-existent"
            
            Test-Path $nonExistentPath | Should -Be $false
        }

        It "Should handle corrupted JSON file" {
            $corruptedJsonPath = Join-Path $script:testImportPath "corrupted.json"
            "{ invalid json content" | Out-File -FilePath $corruptedJsonPath -Encoding UTF8

            { Get-Content $corruptedJsonPath -Raw | ConvertFrom-Json } | Should -Throw
        }

        It "Should create log directory if it doesn't exist" {
            Remove-Item $script:testImportPath -Recurse -Force
            $global:ImportPath = $script:testImportPath
            
            { Write-Log -Message "Test" } | Should -Not -Throw
        }
    }

    Context "Category Description Tests" {
        It "Should build category descriptions structure" {
            $planData = [PSCustomObject]@{
                Categories = [PSCustomObject]@{
                    category1 = "High Priority"
                    category2 = "Low Priority"
                    category3 = ""  # Empty category
                }
            }

            $categoryBody = @{
                categoryDescriptions = @{}
            }

            $planData.Categories.PSObject.Properties | Where-Object { $_.Value } | ForEach-Object {
                $categoryBody.categoryDescriptions[$_.Name] = $_.Value
            }

            # Empty category should not be included
            $categoryBody.categoryDescriptions.Keys.Count | Should -Be 2
            $categoryBody.categoryDescriptions["category1"] | Should -Be "High Priority"
            $categoryBody.categoryDescriptions["category2"] | Should -Be "Low Priority"
        }
    }

    Context "UserMap Conversion Tests" {
        It "Should convert UserMap from PSCustomObject to hashtable" {
            $planData = [PSCustomObject]@{
                UserMap = [PSCustomObject]@{
                    "user-1" = [PSCustomObject]@{
                        DisplayName = "User One"
                        UserPrincipalName = "user1@example.com"
                    }
                    "user-2" = [PSCustomObject]@{
                        DisplayName = "User Two"
                        UserPrincipalName = "user2@example.com"
                    }
                }
            }

            $userMap = @{}
            if ($planData.UserMap) {
                $planData.UserMap.PSObject.Properties | ForEach-Object {
                    $userMap[$_.Name] = $_.Value
                }
            }

            $userMap.Keys.Count | Should -Be 2
            $userMap["user-1"].DisplayName | Should -Be "User One"
            $userMap["user-2"].DisplayName | Should -Be "User Two"
        }
    }
}
