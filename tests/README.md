# Tests for Microsoft Planner Export/Import Tool

This directory contains unit and integration tests for the Planner Export/Import PowerShell scripts using the Pester testing framework.

## Prerequisites

### Install Pester 5.x

```powershell
Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0
```

### Verify Installation

```powershell
Get-Module -ListAvailable Pester
```

## Running Tests

### Run All Tests

```powershell
# From the repository root
Invoke-Pester -Path ./tests
```

### Run Specific Test File

```powershell
# Test Export script only
Invoke-Pester -Path ./tests/Export-PlannerData.Tests.ps1

# Test Import script only
Invoke-Pester -Path ./tests/Import-PlannerData.Tests.ps1
```

### Run Tests with Code Coverage

```powershell
$config = New-PesterConfiguration
$config.Run.Path = './tests'
$config.CodeCoverage.Enabled = $true
$config.CodeCoverage.Path = @(
    './Export-PlannerData.ps1',
    './Import-PlannerData.ps1'
)
$config.Output.Verbosity = 'Detailed'

Invoke-Pester -Configuration $config
```

### Run Tests with Detailed Output

```powershell
Invoke-Pester -Path ./tests -Output Detailed
```

## Test Structure

### Export-PlannerData.Tests.ps1

Tests for the export functionality including:

- **Write-Log Function Tests**
  - Log level handling (INFO, ERROR, WARN, OK)
  - Timestamp formatting
  - Log file creation and appending

- **Export-ReadableSummary Function Tests**
  - Summary file creation
  - Category listing
  - Bucket and task formatting
  - Task completion status
  - Task details (descriptions, checklists, references)
  - Statistics calculation

- **Plan File Name Sanitization Tests**
  - Special character replacement
  - Safe filename generation

- **Data Structure Tests**
  - Export data structure validation
  - Empty plan handling

- **Error Handling Tests**
  - Invalid data structure handling
  - Directory creation

- **User ID Extraction Tests**
  - Assignment user ID extraction
  - Creator ID extraction

- **Priority and Status Mapping Tests**
  - Priority value to text mapping
  - Task completion status mapping

### Import-PlannerData.Tests.ps1

Tests for the import functionality including:

- **Write-Log Function Tests**
  - Log level handling (INFO, ERROR, WARN, DRYRUN)
  - Timestamp formatting
  - Log file creation and appending

- **Resolve-UserId Function Tests**
  - UserMapping usage
  - UPN-based user lookup
  - Fallback mechanisms
  - Error handling for missing users

- **DryRun Mode Tests**
  - No actual resource creation in dry run
  - Dry run logging verification

- **Plan Import Data Structure Tests**
  - JSON file loading
  - Target group resolution
  - Missing data handling

- **Bucket Mapping Tests**
  - Old-to-new bucket ID mapping
  - Task bucket assignment

- **Task Mapping Tests**
  - Old-to-new task ID mapping
  - Completed task filtering

- **Task Body Construction Tests**
  - Basic task properties
  - Optional fields (due date, start date)
  - Category assignment

- **Assignment Tests**
  - Assignment skipping logic
  - Assignment structure creation

- **Task Details Tests**
  - Description handling
  - Checklist structure
  - References/links structure

- **Import Mapping Tests**
  - Mapping file structure
  - Mapping file persistence

- **Error Handling Tests**
  - Missing directory handling
  - Corrupted JSON handling

- **Category and UserMap Tests**
  - Category descriptions
  - UserMap conversion

## Test Coverage

The tests cover the following areas:

| Area | Coverage |
|------|----------|
| Logging Functions | ✅ Comprehensive |
| Data Structures | ✅ Comprehensive |
| File I/O | ✅ Comprehensive |
| User Mapping | ✅ Comprehensive |
| Error Handling | ✅ Comprehensive |
| Business Logic | ✅ Comprehensive |

## Mocking Strategy

These tests focus on **unit testing** the isolated functions without making actual Microsoft Graph API calls. For the functions that interact with the Graph API (like `Connect-ToGraph`, `Get-AllUserPlans`, `Export-PlanDetails`, `Import-PlanFromJson`), we:

1. Test the data structures they work with
2. Test the helper functions in isolation
3. Test error handling logic

To test the full integration with Microsoft Graph API:
- Use a test tenant
- Run the actual scripts with `-DryRun` mode
- Manually verify outputs

## CI/CD Integration

### GitHub Actions Example

Create `.github/workflows/test.yml`:

```yaml
name: Run Tests

on:
  push:
    branches: [ main, develop ]
  pull_request:
    branches: [ main ]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Install Pester
        shell: pwsh
        run: |
          Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0
      
      - name: Run Tests
        shell: pwsh
        run: |
          Invoke-Pester -Path ./tests -Output Detailed -CI
      
      - name: Upload Test Results
        if: always()
        uses: actions/upload-artifact@v3
        with:
          name: test-results
          path: testResults.xml
```

## Adding New Tests

When adding new functionality to the scripts:

1. Add corresponding test cases in the appropriate test file
2. Follow the existing naming conventions:
   - Use descriptive `It` block names
   - Group related tests in `Context` blocks
   - Use `BeforeEach` for test setup
3. Test both happy path and error cases
4. Update this README with new test coverage information

## Test Philosophy

These tests follow these principles:

- **Fast**: No external dependencies, no network calls
- **Isolated**: Each test is independent
- **Repeatable**: Same results every run
- **Readable**: Clear test names and structure
- **Maintainable**: Easy to update when code changes

## Troubleshooting

### Pester Version Issues

If you encounter errors about Pester version conflicts:

```powershell
# Remove old versions
Get-Module Pester -ListAvailable | Where-Object Version -lt '5.0.0' | Uninstall-Module -Force

# Install latest version
Install-Module -Name Pester -Force -SkipPublisherCheck
```

### Import Module Errors

If functions are not being loaded for testing:

1. Ensure the script paths are correct
2. Check that functions are defined with `function` keyword
3. Verify function extraction regex is working

### Test Drive Access Issues

On some systems, `$TestDrive` may have permission issues. If tests fail:

```powershell
# Run PowerShell as Administrator
# Or use a different temp directory
```

## Contributing

When contributing tests:

1. Ensure all tests pass before submitting PR
2. Add tests for new functionality
3. Maintain test coverage above 80%
4. Follow the existing test structure and naming conventions

## License

These tests are part of the Microsoft Planner Export/Import Tool project and follow the same MIT license.
