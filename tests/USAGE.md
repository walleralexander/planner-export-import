# Test Execution Examples

This document provides practical examples of how to run the tests in various scenarios.

## Quick Start

### Run All Tests (Simplest)

```powershell
# From repository root
Invoke-Pester -Path ./tests
```

### Run Tests with Test Runner (Recommended)

```powershell
# From repository root
pwsh ./tests/Run-Tests.ps1

# Or from tests directory
cd tests
pwsh ./Run-Tests.ps1
```

## Test Scenarios

### Scenario 1: Quick Test During Development

When making changes to the scripts, quickly verify your changes:

```powershell
# Run only Export tests
Invoke-Pester -Path ./tests/Export-PlannerData.Tests.ps1

# Run only Import tests  
Invoke-Pester -Path ./tests/Import-PlannerData.Tests.ps1
```

### Scenario 2: Detailed Output for Debugging

When a test fails and you need more information:

```powershell
# Detailed output
pwsh ./tests/Run-Tests.ps1 -Detailed

# Or with Invoke-Pester directly
Invoke-Pester -Path ./tests -Output Detailed
```

### Scenario 3: Run Specific Tests by Name

When you want to run only specific tests:

```powershell
# Run tests matching a pattern
pwsh ./tests/Run-Tests.ps1 -TestName "*Write-Log*"

# Run tests in specific context
pwsh ./tests/Run-Tests.ps1 -TestName "*User ID Extraction*"
```

### Scenario 4: Code Coverage Analysis

To see which parts of your code are tested:

```powershell
# Run with coverage
pwsh ./tests/Run-Tests.ps1 -Coverage

# Coverage report will be saved to coverage.xml
# View in coverage tools or manually inspect
```

### Scenario 5: CI/CD Pipeline

For continuous integration:

```powershell
# CI mode - generates test results XML
pwsh ./tests/Run-Tests.ps1 -CI

# Results saved to testResults.xml (NUnit format)
# Can be consumed by Azure DevOps, Jenkins, GitHub Actions, etc.
```

### Scenario 6: Pre-Commit Verification

Before committing changes:

```powershell
# Quick verification
pwsh ./tests/Run-Tests.ps1

# Exit code will be 0 if all pass, 1 if any fail
# Perfect for git pre-commit hooks
```

## Test Output Examples

### Successful Test Run

```
============================================================
  Microsoft Planner Export/Import Tool - Test Runner
============================================================

Using Pester version: 5.7.1
Test directory: /path/to/tests

Running tests...

Tests completed in 1.85s
Tests Passed: 59, Failed: 0, Skipped: 0, Inconclusive: 0, NotRun: 0

============================================================
  Test Summary
============================================================

Total Tests:   59
Passed:        59
Failed:        0

Duration:      00:00:01.8500000

All tests PASSED!
```

### Failed Test Run

```
Tests completed in 2.1s
Tests Passed: 57, Failed: 2, Skipped: 0, Inconclusive: 0, NotRun: 0

============================================================
  Test Summary
============================================================

Total Tests:   59
Passed:        57
Failed:        2

Duration:      00:00:02.1000000

Tests FAILED!
```

## Understanding Test Results

### Test Categories

The test suite contains 59 tests across 2 main test files:

**Export-PlannerData.Tests.ps1** (21 tests):
- Write-Log function: 5 tests
- Export-ReadableSummary function: 7 tests
- File name sanitization: 2 tests
- Data structures: 2 tests
- Error handling: 2 tests
- User ID extraction: 1 test
- Priority/status mapping: 2 tests

**Import-PlannerData.Tests.ps1** (38 tests):
- Write-Log function: 5 tests
- Resolve-UserId function: 5 tests
- DryRun mode: 2 tests
- Plan import data structures: 3 tests
- Bucket mapping: 3 tests
- Task mapping: 3 tests
- Task body construction: 4 tests
- Assignment tests: 3 tests
- Task details: 3 tests
- Import mapping: 2 tests
- Error handling: 3 tests
- Category/UserMap: 2 tests

## Troubleshooting

### Issue: "Pester not found"

```powershell
# Install Pester
Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0 -Scope CurrentUser

# Verify installation
Get-Module -ListAvailable Pester
```

### Issue: "Cannot resolve parameter set"

This usually means you're using conflicting Pester parameters. Use the test runner instead:

```powershell
# Instead of complex Invoke-Pester commands, use:
pwsh ./tests/Run-Tests.ps1 -Detailed
```

### Issue: Tests fail on first run but pass on second

This can happen if log files from previous tests aren't cleaned up. The tests now handle this automatically by clearing log files at the start of each test.

### Issue: "Access denied" or permission errors

Run PowerShell as Administrator, or change the test output directory:

```powershell
# Tests use $TestDrive which should work without admin rights
# But if issues persist, check your antivirus settings
```

## Best Practices

1. **Always run tests before committing**: Ensure your changes don't break existing functionality

2. **Run tests after pulling changes**: Verify the codebase is in a good state before starting work

3. **Use detailed output when debugging**: It provides much more context when tests fail

4. **Keep tests fast**: These unit tests run in ~2 seconds, which is ideal

5. **Don't modify test data in production**: Always use test tenants for integration tests

6. **Update tests when adding features**: New functionality should have corresponding tests

## Integration with Git Hooks

### Pre-Commit Hook

Create `.git/hooks/pre-commit`:

```bash
#!/bin/bash
echo "Running tests before commit..."
pwsh ./tests/Run-Tests.ps1
exit $?
```

```powershell
# Make executable (Linux/Mac)
chmod +x .git/hooks/pre-commit
```

### Pre-Push Hook

Create `.git/hooks/pre-push`:

```bash
#!/bin/bash
echo "Running full test suite before push..."
pwsh ./tests/Run-Tests.ps1 -Detailed
exit $?
```

## CI/CD Integration Examples

### GitHub Actions

```yaml
name: Tests
on: [push, pull_request]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - name: Install Pester
        shell: pwsh
        run: Install-Module -Name Pester -Force -SkipPublisherCheck
      - name: Run Tests
        shell: pwsh
        run: ./tests/Run-Tests.ps1 -CI
      - name: Upload Results
        if: always()
        uses: actions/upload-artifact@v3
        with:
          name: test-results
          path: testResults.xml
```

### Azure DevOps

```yaml
steps:
- pwsh: |
    Install-Module -Name Pester -Force -SkipPublisherCheck
    ./tests/Run-Tests.ps1 -CI
  displayName: 'Run Tests'
  
- task: PublishTestResults@2
  inputs:
    testResultsFormat: 'NUnit'
    testResultsFiles: '**/testResults.xml'
  condition: always()
```

## Next Steps

After running unit tests successfully:

1. Review `Integration-Tests.ps1` for manual test scenarios
2. Set up a test Microsoft 365 tenant
3. Run actual export/import with test data
4. Verify results manually in Planner web interface
5. Document any issues or edge cases found

## Support

For questions or issues with tests:
- Check the main README.md for prerequisites
- Review test code for examples of proper usage
- Open an issue on GitHub with test output if tests fail unexpectedly
