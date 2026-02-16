# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a PowerShell-based Microsoft Planner export/import tool created by Alexander Waller. It enables complete backup and restoration of Microsoft Planner data via Microsoft Graph API, primarily designed for license migration scenarios.

**Key capabilities:**

- Exports all Planner data (plans, buckets, tasks, checklists, descriptions, assignments, categories, references)
- Imports data to same or different Microsoft 365 groups
- Handles user mapping for cross-tenant migrations
- Includes rate limiting, retry logic, and dry-run mode

## Commands

### Export Operations

```powershell
# User-based: Export all plans of current user
.\Export-PlannerData.ps1 -UseCurrentUser

# User-based: Export with completed tasks
.\Export-PlannerData.ps1 -UseCurrentUser -IncludeCompletedTasks

# Group-based: Interactive selection of M365 groups/SharePoint sites
.\Export-PlannerData.ps1 -Interactive

# Group-based: Export from specific groups by name
.\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha", "Marketing Team"

# Group-based: Export from specific groups by ID
.\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."

# Export to specific directory
.\Export-PlannerData.ps1 -GroupNames "Projektteam" -ExportPath "C:\Backup\Planner"

# Include completed tasks in export
.\Export-PlannerData.ps1 -GroupNames "Projektteam" -IncludeCompletedTasks

# Note: Without parameters, the script will prompt for group selection
```

### Import Operations

```powershell
# Import all plans (to original groups)
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000"

# Dry run - preview what would be imported
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -DryRun

# Import to specific target group
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -TargetGroupId "neue-gruppe-id"

# Skip assignments during import
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipAssignments

# Skip completed tasks
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipCompletedTasks

# Import with user mapping (for tenant migrations)
$mapping = @{ "alte-user-id-1" = "neue-user-id-1"; "alte-user-id-2" = "neue-user-id-2" }
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -UserMapping $mapping
```

### Testing/Debugging

```powershell
# Test connection and permissions
Connect-MgGraph -Scopes "Group.Read.All", "Tasks.Read", "User.Read"
Get-MgContext

# Check what groups the current user belongs to
Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me/memberOf/microsoft.graph.group"

# View exported data structure
Get-Content ".\PlannerExport_YYYYMMDD_HHMMSS\PlanName.json" | ConvertFrom-Json | Format-List
```

## Architecture

### Script Structure

Both scripts follow the same organizational pattern:

1. **Parameter block** - Script parameters with validation
2. **Functions region** (`#region Funktionen`) - All helper functions
3. **Main program region** (`#region Hauptprogramm`) - Main execution flow

### Key Functions

**Export-PlannerData.ps1:**

- `Connect-ToGraph` - Authenticates with Microsoft Graph (scopes: Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read, User.ReadBasic.All)
- `Get-AllUserPlans` - Retrieves all plans the current user has access to (via user's groups and /me/planner/plans endpoint)
- `Get-AllM365Groups` - Retrieves all M365 groups (Unified groups) with paging support
- `Get-GroupsByNames` - Searches for M365 groups by display name (supports partial matching)
- `Show-GroupSelectionMenu` - Interactive menu for selecting groups from a list
- `Get-PlansByGroupIds` - Retrieves plans for specific group IDs
- `Export-PlanDetails` - Main export logic: loads plan, buckets, tasks, task details, user info; handles paging
- `Export-ReadableSummary` - Creates human-readable text summary of plan data
- `Write-PlannerLog` - Unified logging to console and file

**Import-PlannerData.ps1:**

- `Connect-ToGraph` - Authenticates with elevated permissions (Group.ReadWrite.All, Tasks.ReadWrite)
- `Invoke-GraphWithRetry` - Handles API rate limiting (429 errors), retry logic with exponential backoff
- `Resolve-UserId` - Maps old user IDs to new tenant using UserMapping, UPN lookup, or direct ID validation
- `Import-PlanFromJson` - Main import logic: creates plan, sets categories, creates buckets/tasks, sets task details (descriptions, checklists, references)
- `Write-Log` - Unified logging with DryRun mode support

### Microsoft Graph API Usage

**Critical implementation details:**

1. **ETag handling**: PATCH operations require If-Match header with current ETag
   - Plan details: GET first to retrieve ETag, then PATCH with `If-Match: <etag>`
   - Task details: Same pattern

2. **Rate limiting**:
   - Built-in throttle delay: `$ThrottleDelayMs` (default 500ms between requests)
   - Automatic retry on 429 responses with Retry-After header parsing
   - Task detail loading includes 200ms sleep between requests

3. **Paging support**: Export handles `@odata.nextLink` for large task collections

4. **User resolution strategy** (Import):
   - First: Check UserMapping hashtable
   - Second: Lookup by UserPrincipalName in new tenant
   - Third: Lookup by Mail if different from UPN
   - Fallback: Try original ID directly (same tenant scenario)

### Data Model

**Export output structure:**

```Text
PlannerExport_YYYYMMDD_HHMMSS/
├── _ExportIndex.json                 # Metadata: export date, user, plan count
├── export.log                        # Timestamped execution log
├── <PlanName>.json                   # Full structured data
│   ├── Plan                          # Plan object with groupId, groupDisplayName
│   ├── Buckets                       # Array of bucket objects
│   ├── Tasks                         # Array of task objects
│   ├── TaskDetails                   # Array with taskId references
│   ├── Categories                    # categoryDescriptions from plan details
│   ├── PlanDetails                   # Full plan details object
│   └── UserMap                       # userId -> {DisplayName, UPN, Mail}
└── <PlanName>_Zusammenfassung.txt   # Human-readable summary
```

**Import mapping output:**

```json
{
  "ImportDate": "ISO8601",
  "OriginalPlan": "original-plan-id",
  "NewPlanId": "new-plan-id",
  "GroupId": "target-group-id",
  "BucketMap": { "old-id": "new-id" },
  "TaskMap": { "old-task-id": "new-task-id" }
}
```

### Known Limitations

- **Comments** cannot be exported/imported (stored in Exchange, not accessible via Planner API)
- **File attachments** are exported as link references only (actual files remain in SharePoint)
- **Task history** (audit trail) is not exported
- **Rate limits**: ~2000 requests/minute per Microsoft Graph throttling policies

## Prerequisites

- PowerShell 7+ recommended (minimum: PowerShell 5.1)
- Microsoft.Graph PowerShell module: `Install-Module Microsoft.Graph -Scope CurrentUser`
- Appropriate Microsoft 365 licenses (Planner access)
- Admin consent may be required for Group.ReadWrite.All scope

## Important Notes

- All log output is in German
- Scripts use UTF-8 encoding for all file operations
- Plan titles with special characters are sanitized for filenames: `[\\/:*?"<>|]` replaced with `_`
- Assignment restoration only works if users exist in target tenant with matching UPN/Mail
- Always test imports with `-DryRun` first to preview changes
