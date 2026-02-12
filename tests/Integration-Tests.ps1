<#
.SYNOPSIS
    Integration test scenarios for Microsoft Planner Export/Import Tool

.DESCRIPTION
    This file documents integration test scenarios that should be performed
    manually or in a test environment with actual Microsoft 365 tenants.
    
    These tests require real Microsoft Graph API access and cannot be
    automated without a test tenant.

.NOTES
    Author: Alexander Waller
    Date: 2026-02-12
#>

<#
===============================================================================
INTEGRATION TEST SCENARIOS
===============================================================================

These scenarios should be tested in a real Microsoft 365 environment with
appropriate test data. Always use a test tenant, never production data.

PREREQUISITES:
- Test Microsoft 365 tenant
- Test user accounts
- Test Microsoft 365 groups with Planner
- Microsoft.Graph PowerShell module installed
- Appropriate permissions (see README.md)

===============================================================================
#>

# Scenario 1: Export Single Plan
<#
TEST: Export a single plan from a known group

STEPS:
1. Create a test group with a simple plan containing:
   - 2-3 buckets
   - 5-10 tasks
   - Tasks with descriptions, checklists, due dates
   - At least one completed task
   - At least one task with assignments

2. Run export:
   .\Export-PlannerData.ps1 -GroupIds "YOUR-GROUP-ID" -ExportPath ".\test-export-single"

EXPECTED RESULTS:
- Export completes without errors
- Directory created: test-export-single
- Files created:
  * PlanName.json
  * PlanName_Zusammenfassung.txt
  * _ExportIndex.json
  * export.log
- JSON file contains all expected data structures
- Summary file is human-readable and accurate
- Log file shows no errors

VALIDATION:
- Open PlanName_Zusammenfassung.txt and verify:
  * All buckets are listed
  * All tasks are listed with correct details
  * Task counts match
  * Assignments are listed
- Open PlanName.json and verify structure:
  * Plan, Buckets, Tasks, TaskDetails, Categories, UserMap all present
#>

# Scenario 2: Export All User Plans
<#
TEST: Export all plans accessible to the current user

STEPS:
1. Ensure user is member of 2-3 groups with plans
2. Run export:
   .\Export-PlannerData.ps1 -ExportPath ".\test-export-all"

EXPECTED RESULTS:
- All accessible plans are discovered
- Multiple JSON and summary files created
- _ExportIndex.json lists all plans
- No errors in export.log

VALIDATION:
- Count of exported plans matches expected count
- All plans are accessible in the original Planner
- File sizes are reasonable (not empty)
#>

# Scenario 3: Import to Same Group (Restoration)
<#
TEST: Import a previously exported plan back to its original group

STEPS:
1. Export a test plan (see Scenario 1)
2. Manually delete the plan from Planner
3. Run import:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single"

EXPECTED RESULTS:
- Import completes successfully
- New plan created with same title
- All buckets recreated
- All tasks recreated with correct properties
- Task details (descriptions, checklists) restored
- Categories/labels restored
- No errors in import.log

VALIDATION:
- Open restored plan in Planner web interface
- Verify buckets and tasks match summary file
- Verify task descriptions and checklists
- Verify due dates and priorities
- Verify categories are correct
#>

# Scenario 4: Import to Different Group
<#
TEST: Import a plan to a different Microsoft 365 group

STEPS:
1. Export a plan from Group A
2. Create a new empty Group B
3. Run import:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single" -TargetGroupId "GROUP-B-ID"

EXPECTED RESULTS:
- Plan created in Group B
- All data transferred correctly
- Import mapping file created showing old->new IDs

VALIDATION:
- Verify plan appears in Group B
- Verify all tasks and buckets present
- Verify no data loss
- Check ImportMapping.json for correct mappings
#>

# Scenario 5: Dry Run Mode
<#
TEST: Verify dry run mode doesn't make actual changes

STEPS:
1. Have an export ready
2. Run import in dry run mode:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single" -DryRun

EXPECTED RESULTS:
- Script runs without errors
- Log shows [DRYRUN] entries
- No actual plans/tasks created in target group
- Preview of what would be imported is shown

VALIDATION:
- Check target group - should be unchanged
- Review import.log for DRYRUN messages
- Verify counts shown in output
#>

# Scenario 6: Skip Completed Tasks
<#
TEST: Import without completed tasks

STEPS:
1. Export a plan containing completed and incomplete tasks
2. Run import with skip flag:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single" -SkipCompletedTasks

EXPECTED RESULTS:
- Only incomplete tasks are imported
- Completed tasks are skipped
- Log shows which tasks were skipped

VALIDATION:
- Count of imported tasks should be less than exported
- No completed tasks in new plan
- Incomplete tasks all present
#>

# Scenario 7: Skip Assignments
<#
TEST: Import without user assignments

STEPS:
1. Export a plan with task assignments
2. Run import without assignments:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single" -SkipAssignments

EXPECTED RESULTS:
- Tasks created but unassigned
- No errors about missing users
- All other task properties preserved

VALIDATION:
- Tasks exist but have no assignments
- Other task data intact
#>

# Scenario 8: User Mapping (Cross-Tenant)
<#
TEST: Import with user mapping for different tenant

NOTE: This requires two tenants to test properly

STEPS:
1. Export from Tenant A
2. Create user mapping:
   $mapping = @{
       "old-user-id-from-tenantA" = "new-user-id-in-tenantB"
   }
3. Run import in Tenant B:
   .\Import-PlannerData.ps1 -ImportPath ".\test-export-single" -UserMapping $mapping

EXPECTED RESULTS:
- Tasks assigned to mapped users in new tenant
- Unmapped users logged as warnings

VALIDATION:
- Check assignments in new plan
- Verify mapped users are correctly assigned
- Check log for warnings about unmapped users
#>

# Scenario 9: Large Plan (Performance)
<#
TEST: Export and import a large plan

STEPS:
1. Create or use a plan with:
   - 10+ buckets
   - 100+ tasks
   - Many task details (checklists, descriptions)
2. Export:
   .\Export-PlannerData.ps1 -GroupIds "GROUP-ID" -ExportPath ".\test-large"
3. Import:
   .\Import-PlannerData.ps1 -ImportPath ".\test-large"

EXPECTED RESULTS:
- Export completes in reasonable time (consider API throttling)
- Import completes without rate limit errors
- All data transferred correctly
- Progress indicators work properly

VALIDATION:
- Compare task counts before/after
- Spot-check random tasks for data integrity
- Review logs for any throttling messages
#>

# Scenario 10: Special Characters in Plan Names
<#
TEST: Handle plan names with special characters

STEPS:
1. Create a plan with special characters in title:
   "Plan: Test/Update & Review [2026] (Q1)"
2. Export the plan
3. Import the plan

EXPECTED RESULTS:
- File name sanitized: "Plan__Test_Update___Review__2026___Q1_.json"
- Export and import both succeed
- Plan title preserved correctly in import

VALIDATION:
- Check file name is valid (no illegal characters)
- Verify plan title in imported plan matches original
- No file system errors
#>

# Scenario 11: Plans with Attachments/References
<#
TEST: Handle task references and links

STEPS:
1. Create tasks with:
   - SharePoint document links
   - External URL references
   - Multiple references per task
2. Export and import

EXPECTED RESULTS:
- References exported in TaskDetails
- References recreated in imported tasks
- Links are functional in imported plan

VALIDATION:
- Check references section in JSON
- Verify links appear in imported tasks
- Click links to verify they work
#>

# Scenario 12: Error Recovery
<#
TEST: Handle errors gracefully during import

STEPS:
1. Export a plan
2. Modify JSON to introduce errors:
   - Invalid group ID
   - Missing required fields
   - Corrupted data
3. Attempt import

EXPECTED RESULTS:
- Script detects errors
- Error messages are clear and helpful
- No partial imports (if possible)
- Log file contains error details

VALIDATION:
- Review error messages
- Check log file for details
- Verify system is in consistent state
#>

# Scenario 13: Rate Limiting Handling
<#
TEST: Verify rate limiting is handled correctly

STEPS:
1. Export multiple large plans rapidly
2. Or import multiple plans in succession

EXPECTED RESULTS:
- Script automatically waits when rate limited
- 429 errors are caught and handled
- Retries succeed after waiting
- No permanent failures due to rate limiting

VALIDATION:
- Monitor log for rate limit messages
- Verify Retry-After headers are respected
- All operations eventually succeed
#>

# Scenario 14: Include Completed Tasks Flag
<#
TEST: Export with completed tasks explicitly included

STEPS:
1. Create plan with mix of completed and incomplete tasks
2. Export with flag:
   .\Export-PlannerData.ps1 -IncludeCompletedTasks -ExportPath ".\test-with-completed"
3. Export without flag (default):
   .\Export-PlannerData.ps1 -ExportPath ".\test-default"

EXPECTED RESULTS:
- Both exports include all tasks (flag doesn't filter, just notes in log)
- Log messages indicate what was exported

VALIDATION:
- Compare task counts in both exports
- Verify completed tasks in both exports
#>

<#
===============================================================================
MANUAL VALIDATION CHECKLIST
===============================================================================

After running integration tests, verify:

EXPORT VALIDATION:
□ All plans discovered correctly
□ All buckets exported
□ All tasks exported (or filtered correctly)
□ Task details complete (descriptions, checklists, due dates)
□ User information captured
□ Categories/labels exported
□ References/links captured
□ File encoding is UTF-8
□ JSON is valid and well-formed
□ Summary files are readable and accurate
□ Export index is complete
□ Log file shows no unexpected errors

IMPORT VALIDATION:
□ Plans created successfully
□ Buckets recreated in correct order
□ Tasks recreated with all properties
□ Task details restored (descriptions, checklists)
□ Due dates preserved
□ Start dates preserved
□ Priorities correct
□ Categories/labels applied
□ Assignments restored (or skipped correctly)
□ References/links working
□ Import mapping file created
□ No duplicate items created
□ Log file shows detailed progress

CROSS-TENANT MIGRATION:
□ User mapping works correctly
□ Unmapped users logged as warnings
□ All other data transferred correctly

ERROR HANDLING:
□ Invalid data caught and reported
□ Partial failures don't corrupt data
□ Clear error messages
□ Errors logged with details

PERFORMANCE:
□ Large plans handled efficiently
□ Rate limiting handled gracefully
□ Progress indicators accurate
□ Memory usage reasonable

===============================================================================
#>

# Helper function to create test data
function New-TestPlanData {
    param(
        [string]$GroupId,
        [string]$PlanTitle = "Integration Test Plan - $(Get-Date -Format 'yyyyMMdd-HHmmss')",
        [int]$BucketCount = 3,
        [int]$TaskCount = 10
    )
    
    Write-Host "This is a documentation file only." -ForegroundColor Yellow
    Write-Host "Use the Microsoft Planner web interface or Graph API to create test data." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Recommended test data structure:" -ForegroundColor Cyan
    Write-Host "  - $BucketCount buckets (e.g., 'To Do', 'In Progress', 'Done')" -ForegroundColor Gray
    Write-Host "  - $TaskCount tasks distributed across buckets" -ForegroundColor Gray
    Write-Host "  - Mix of completed and incomplete tasks" -ForegroundColor Gray
    Write-Host "  - Tasks with descriptions, checklists, and due dates" -ForegroundColor Gray
    Write-Host "  - At least 2 users assigned to tasks" -ForegroundColor Gray
    Write-Host "  - 2-3 categories/labels defined and used" -ForegroundColor Gray
    Write-Host "  - Some tasks with external references/links" -ForegroundColor Gray
}

Write-Host @"

================================================================================
  INTEGRATION TEST DOCUMENTATION
================================================================================

This file documents manual integration test scenarios for the Microsoft
Planner Export/Import Tool.

These scenarios require a real Microsoft 365 environment and should be
performed before releasing new versions or after significant changes.

To run integration tests:
1. Set up a test Microsoft 365 tenant (never use production!)
2. Create test groups and plans with diverse data
3. Follow the scenarios documented above
4. Use the validation checklists
5. Document any issues found

For automated unit tests, use:
  .\tests\Run-Tests.ps1

For more information, see:
  .\tests\README.md

================================================================================

"@ -ForegroundColor Cyan
