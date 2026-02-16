# Microsoft Planner Export/Import Tool - Copilot Instructions

## Project Overview

This is a PowerShell-based Microsoft Planner export/import tool created by Alexander Waller for backing up and restoring Microsoft Planner data via Microsoft Graph API. It's primarily designed for license migration scenarios where organizations need to preserve their Planner data when switching Microsoft 365 licenses or tenants.

**Project Type**: PowerShell scripts (2 main files)
**Primary Language**: PowerShell 7+ (minimum 5.1)
**API Integration**: Microsoft Graph API v1.0
**Testing Framework**: Pester 5.x
**Repository Size**: Small (~50KB of PowerShell code, extensive documentation)

**Key Capabilities**:
- Export all Planner data (plans, buckets, tasks, checklists, descriptions, assignments, categories, references)
- Import data to same or different Microsoft 365 groups
- Handle user mapping for cross-tenant migrations
- Rate limiting, retry logic, and dry-run mode

## Testing

### Running Tests

**ALWAYS run tests using Pester before making changes to understand any pre-existing failures:**

```powershell
# Run all tests from repository root
Invoke-Pester -Path ./tests -Output Detailed

# Run specific test file
Invoke-Pester -Path ./tests/Export-PlannerData.Tests.ps1
Invoke-Pester -Path ./tests/Import-PlannerData.Tests.ps1

# Run with code coverage
$config = New-PesterConfiguration
$config.Run.Path = './tests'
$config.CodeCoverage.Enabled = $true
$config.CodeCoverage.Path = @('./Export-PlannerData.ps1', './Import-PlannerData.ps1')
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
```

**Test Infrastructure**:
- **Total Tests**: 97 tests (93 passing, 4 pre-existing failures)
- **Export Tests**: 59 tests in `tests/Export-PlannerData.Tests.ps1`
- **Import Tests**: 38 tests in `tests/Import-PlannerData.Tests.ps1`
- **Integration Tests**: Manual scenarios in `tests/Integration-Tests.ps1`
- **Prerequisites**: Install Pester 5.x with `Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0`

**Important**: Tests are unit tests that don't make real API calls. They use mocking to test function logic in isolation. For full integration testing, use a test Microsoft 365 tenant.

## Script Structure and Architecture

Both `Export-PlannerData.ps1` and `Import-PlannerData.ps1` follow the same organizational pattern:

1. **Parameter block** - Script parameters with validation attributes
2. **Functions region** (`#region Funktionen`) - All helper functions
3. **Main program region** (`#region Hauptprogramm`) - Main execution flow

### Export-PlannerData.ps1 Key Functions

- `Connect-ToGraph` - Authenticates with Microsoft Graph (scopes: Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read, User.ReadBasic.All)
- `Get-AllUserPlans` - Retrieves all plans the current user has access to via user's groups and /me/planner/plans endpoint
- `Get-AllM365Groups` - Retrieves all M365 groups (Unified groups) with paging support
- `Get-GroupsByNames` - Searches for M365 groups by display name (supports partial matching)
- `Show-GroupSelectionMenu` - Interactive menu for selecting groups from a list
- `Get-PlansByGroupIds` - Retrieves plans for specific group IDs
- `Export-PlanDetails` - Main export logic: loads plan, buckets, tasks, task details, user info; handles paging
- `Export-ReadableSummary` - Creates human-readable text summary of plan data
- `Write-PlannerLog` - Unified logging to console and file

### Import-PlannerData.ps1 Key Functions

- `Connect-ToGraph` - Authenticates with elevated permissions (Group.ReadWrite.All, Tasks.ReadWrite)
- `Invoke-GraphWithRetry` - Handles API rate limiting (429 errors), retry logic with exponential backoff
- `Resolve-UserId` - Maps old user IDs to new tenant using UserMapping, UPN lookup, or direct ID validation
- `Import-PlanFromJson` - Main import logic: creates plan, sets categories, creates buckets/tasks, sets task details
- `Write-Log` - Unified logging with DryRun mode support

## Common Commands

### Export Operations

```powershell
# User-based: Export all plans of current user
.\Export-PlannerData.ps1 -UseCurrentUser

# User-based: Export with completed tasks
.\Export-PlannerData.ps1 -UseCurrentUser -IncludeCompletedTasks

# Group-based: Interactive selection
.\Export-PlannerData.ps1 -Interactive

# Group-based: Export from specific groups by name
.\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha", "Marketing Team"

# Group-based: Export from specific groups by ID
.\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."

# Export to specific directory
.\Export-PlannerData.ps1 -GroupNames "Projektteam" -ExportPath "C:\Backup\Planner"
```

### Import Operations

```powershell
# Import all plans to original groups
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000"

# Dry run - preview what would be imported (ALWAYS do this first)
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -DryRun

# Import to specific target group
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -TargetGroupId "neue-gruppe-id"

# Skip assignments during import
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipAssignments

# Import with user mapping (for tenant migrations)
$mapping = @{ "alte-user-id-1" = "neue-user-id-1"; "alte-user-id-2" = "neue-user-id-2" }
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -UserMapping $mapping
```

## Critical Implementation Details

### Microsoft Graph API Best Practices

**1. ETag Handling** - MANDATORY for all PATCH operations:
- Plan details: GET first to retrieve ETag, then PATCH with `If-Match: <etag>` header
- Task details: Same pattern - always GET before PATCH
- **Never skip ETag headers** - operations will fail with 412 Precondition Failed

**2. Rate Limiting Strategy**:
- Built-in throttle delay: `$ThrottleDelayMs` parameter (default 500ms between requests)
- Automatic retry on 429 responses with Retry-After header parsing in `Invoke-GraphWithRetry`
- Task detail loading includes 200ms sleep between requests
- **Microsoft Graph limit**: ~2000 requests/minute per user/app

**3. Paging Support**:
- Export handles `@odata.nextLink` for large task collections
- Always check for `@odata.nextLink` in responses and follow it
- M365 groups API also requires paging support

**4. User Resolution Strategy (Import)**:
- First: Check UserMapping hashtable (explicit mapping provided by user)
- Second: Lookup by UserPrincipalName in new tenant
- Third: Lookup by Mail if different from UPN
- Fallback: Try original ID directly (same tenant scenario)
- **Always validate user exists before assignment**

### Data Model

**Export Output Structure:**
```
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

**Import Mapping Output** (`<PlanName>_ImportMapping.json`):
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

## Important Code Conventions

### File Handling
- **Always use UTF-8 encoding** for all file operations
- Plan titles with special characters are sanitized for filenames: `[\\/:*?"<>|]` replaced with `_`
- Export path defaults to `C:\planner-data\PlannerExport_YYYYMMDD_HHMMSS`

### Logging
- All log output is in **German** (log messages should maintain this convention)
- Log levels: INFO, ERROR, WARN, OK (Export), DRYRUN (Import)
- Timestamps use format: `yyyy-MM-dd HH:mm:ss`
- Both scripts write to `export.log` or `import.log` in their respective directories

### PowerShell Conventions
- Use `#region` and `#endregion` to organize code sections
- Functions should have descriptive names with verb-noun pattern
- Always use `$null` comparison on the left side: `if ($null -eq $variable)`
- Use `-ErrorAction Stop` for API calls that should halt on error
- Prefer `Write-Host` for console output, custom logging functions for file logging

## Known Limitations

- **Comments** cannot be exported/imported (stored in Exchange, not accessible via Planner API)
- **File attachments** are exported as link references only (actual files remain in SharePoint)
- **Task history** (audit trail) is not exported
- **Assignment restoration** only works if users exist in target tenant with matching UPN/Mail

## Prerequisites and Setup

### Required Software
- PowerShell 7+ recommended (minimum: PowerShell 5.1)
- Microsoft.Graph PowerShell module: `Install-Module Microsoft.Graph -Scope CurrentUser`
- Pester 5.x for testing: `Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0`

### Required Permissions
**Export**: Group.Read.All, Tasks.Read, Tasks.ReadWrite, User.Read, User.ReadBasic.All
**Import**: Group.ReadWrite.All, Tasks.ReadWrite, User.Read, User.ReadBasic.All
- Admin consent may be required for Group.ReadWrite.All scope

### Authentication
- Scripts use interactive browser-based authentication via `Connect-MgGraph`
- First run opens browser for sign-in and consent
- Subsequent runs may use cached tokens

## Project Files

**Root Directory**:
- `Export-PlannerData.ps1` - Main export script (~32KB, 900+ lines)
- `Import-PlannerData.ps1` - Main import script (~21KB, 600+ lines)
- `README.md` - User documentation with examples (German/English)
- `CLAUDE.md` - AI agent instructions (comprehensive reference)
- `LICENSE` - MIT License
- `.gitignore` - Excludes logs, export data, temp files

**Tests Directory** (`/tests/`):
- `Export-PlannerData.Tests.ps1` - 59 unit tests for export functionality
- `Import-PlannerData.Tests.ps1` - 38 unit tests for import functionality
- `Integration-Tests.ps1` - Manual integration testing scenarios
- `Run-Tests.ps1` - Test runner with options
- `README.md` - Comprehensive testing documentation
- `USAGE.md` - Practical test examples and CI/CD integration

## Validation Steps

Before committing changes:
1. **Run all tests**: `Invoke-Pester -Path ./tests -Output Detailed`
2. **Check for PowerShell syntax errors**: `Get-Command -Syntax` on modified functions
3. **Verify UTF-8 encoding**: Ensure files are saved with UTF-8 encoding
4. **Test export/import in dry-run mode** if modifying core logic
5. **Review log output** to ensure German language conventions are maintained

## Development Workflow

1. **Before making changes**: Run `Invoke-Pester -Path ./tests` to establish baseline
2. **Make minimal changes**: Modify only what's necessary for the task
3. **Update tests**: Add/modify tests for new functionality in appropriate test file
4. **Verify changes**: Run tests again to ensure no regressions
5. **Test manually** (if possible): Run scripts with `-DryRun` mode to verify behavior
6. **Update documentation**: If changing user-facing behavior, update README.md

## Common Issues and Solutions

| Problem | Solution |
|---------|----------|
| "Insufficient privileges" | Admin must grant app permissions in Azure AD |
| "429 Too Many Requests" | Script handles automatically via `Invoke-GraphWithRetry` |
| Missing assignments after import | Users don't exist in target tenant - use `-SkipAssignments` or UserMapping |
| Pester version conflicts | Uninstall Pester < 5.0.0 and install latest: `Install-Module -Name Pester -Force` |
| ETag errors (412 Precondition Failed) | Always GET resource first to obtain current ETag before PATCH |

## Additional Resources

- **CLAUDE.md**: Comprehensive technical documentation for AI agents
- **tests/README.md**: Detailed testing documentation with coverage information
- **README.md**: User guide with practical examples in German
- **Microsoft Graph API Documentation**: https://learn.microsoft.com/en-us/graph/api/resources/planner-overview
