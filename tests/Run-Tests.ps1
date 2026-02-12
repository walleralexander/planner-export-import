<#
.SYNOPSIS
    Test runner script for Microsoft Planner Export/Import Tool

.DESCRIPTION
    This script provides a convenient way to run all tests with various options.
    It checks for Pester installation, runs tests, and displays results.

.PARAMETER Coverage
    Run tests with code coverage analysis

.PARAMETER Detailed
    Show detailed test output

.PARAMETER CI
    Run in CI mode (for continuous integration)

.PARAMETER TestName
    Run specific test(s) by name (supports wildcards)

.EXAMPLE
    .\Run-Tests.ps1
    Run all tests with standard output

.EXAMPLE
    .\Run-Tests.ps1 -Detailed
    Run all tests with detailed output

.EXAMPLE
    .\Run-Tests.ps1 -Coverage
    Run all tests with code coverage analysis

.EXAMPLE
    .\Run-Tests.ps1 -TestName "*Write-PlannerLog*"
    Run only tests matching "Write-PlannerLog"

.NOTES
    Author: Alexander Waller
    Date: 2026-02-12
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$Coverage,

    [Parameter(Mandatory = $false)]
    [switch]$Detailed,

    [Parameter(Mandatory = $false)]
    [switch]$CI,

    [Parameter(Mandatory = $false)]
    [string]$TestName
)

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Planner Export/Import Tool - Test Runner" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check if Pester is installed
$pesterModule = Get-Module -ListAvailable -Name Pester | 
    Where-Object { $_.Version -ge [Version]'5.0.0' } | 
    Select-Object -First 1

if (-not $pesterModule) {
    Write-Host "Pester 5.x is not installed!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Installing Pester..." -ForegroundColor Yellow
    
    try {
        Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0 -Scope CurrentUser
        Write-Host "Pester installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install Pester: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Please install Pester manually:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name Pester -Force -SkipPublisherCheck -MinimumVersion 5.0.0" -ForegroundColor Gray
        exit 1
    }
}
else {
    Write-Host "Using Pester version: $($pesterModule.Version)" -ForegroundColor Green
}

# Import Pester
Import-Module Pester -MinimumVersion 5.0.0

# Get test directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# If we're already in the tests directory, use current directory
# Otherwise look for tests subdirectory
if ((Split-Path -Leaf $scriptPath) -eq "tests") {
    $testsPath = $scriptPath
}
else {
    $testsPath = Join-Path $scriptPath "tests"
}

if (-not (Test-Path $testsPath)) {
    Write-Host "Tests directory not found: $testsPath" -ForegroundColor Red
    exit 1
}

Write-Host "Test directory: $testsPath" -ForegroundColor Gray
Write-Host ""

# Configure Pester
$config = New-PesterConfiguration

# Set paths
$config.Run.Path = $testsPath

# Set output verbosity
if ($Detailed) {
    $config.Output.Verbosity = 'Detailed'
}
else {
    $config.Output.Verbosity = 'Normal'
}

# Set CI mode
if ($CI) {
    $config.TestResult.Enabled = $true
    $config.TestResult.OutputPath = Join-Path $scriptPath "testResults.xml"
    $config.TestResult.OutputFormat = 'NUnitXml'
}

# Set test name filter
if ($TestName) {
    $config.Filter.FullName = $TestName
    Write-Host "Running tests matching: $TestName" -ForegroundColor Yellow
    Write-Host ""
}

# Configure code coverage
if ($Coverage) {
    Write-Host "Code coverage analysis enabled" -ForegroundColor Yellow
    Write-Host ""
    
    $config.CodeCoverage.Enabled = $true
    $config.CodeCoverage.Path = @(
        (Join-Path $scriptPath "Export-PlannerData.ps1"),
        (Join-Path $scriptPath "Import-PlannerData.ps1")
    )
    $config.CodeCoverage.OutputFormat = 'JaCoCo'
    $config.CodeCoverage.OutputPath = Join-Path $scriptPath "coverage.xml"
}

# Run tests
Write-Host "Running tests..." -ForegroundColor Cyan
Write-Host ""

$config.Run.PassThru = $true
$result = Invoke-Pester -Configuration $config

# Display summary
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Pester 5.x result structure
if ($result) {
    $totalTests = if ($result.TotalCount) { $result.TotalCount } else { 0 }
    $passedTests = if ($result.PassedCount) { $result.PassedCount } else { 0 }
    $failedTests = if ($result.FailedCount) { $result.FailedCount } else { 0 }
    $skippedTests = if ($result.SkippedCount) { $result.SkippedCount } else { 0 }
    
    Write-Host "Total Tests:   $totalTests" -ForegroundColor White
    Write-Host "Passed:        $passedTests" -ForegroundColor Green
    if ($failedTests -gt 0) {
        Write-Host "Failed:        $failedTests" -ForegroundColor Red
    }
    else {
        Write-Host "Failed:        $failedTests" -ForegroundColor Green
    }
    if ($skippedTests -gt 0) {
        Write-Host "Skipped:       $skippedTests" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Duration:      $($result.Duration)" -ForegroundColor Gray
    Write-Host ""
}
else {
    Write-Host "No test results available" -ForegroundColor Red
}

# Code coverage summary
if ($Coverage -and $result.CodeCoverage) {
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  Code Coverage" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $covered = $result.CodeCoverage.CoveredCommands
    $total = $result.CodeCoverage.TotalCommands
    $percentage = if ($total -gt 0) { [math]::Round(($covered / $total) * 100, 2) } else { 0 }
    
    $coverageColor = if ($percentage -ge 80) { "Green" } 
                     elseif ($percentage -ge 60) { "Yellow" } 
                     else { "Red" }
    
    Write-Host "Commands Covered: $covered / $total" -ForegroundColor White
    Write-Host "Coverage:         $percentage%" -ForegroundColor $coverageColor
    Write-Host ""
    Write-Host "Coverage report saved to: coverage.xml" -ForegroundColor Gray
    Write-Host ""
}

# Exit code based on test results
if ($result -and $result.FailedCount -gt 0) {
    Write-Host "Tests FAILED!" -ForegroundColor Red
    exit 1
}
else {
    Write-Host "All tests PASSED!" -ForegroundColor Green
    exit 0
}
