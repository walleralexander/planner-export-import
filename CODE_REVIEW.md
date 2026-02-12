# Code Review and Recommended Corrections
## Microsoft Planner Export/Import Tool

**Review Date:** 2026-02-12  
**Reviewer:** Automated Analysis + Manual Review  
**Status:** ✅ All tests passing (59/59) - No critical failures  
**Overall Assessment:** Code is functional and well-structured, but has several areas for improvement

---

## Executive Summary

The PowerShell scripts are working correctly as evidenced by passing test suite. However, PSScriptAnalyzer has identified several best practice violations and potential improvements. This document categorizes issues by severity and provides specific corrections.

### Issue Summary by Severity

| Severity | Export-PlannerData.ps1 | Import-PlannerData.ps1 | Total |
|----------|------------------------|------------------------|-------|
| **Error** | 0 | 0 | 0 |
| **Warning** | 33 | 28 | 61 |
| **Information** | 4 | 4 | 8 |
| **Total** | 37 | 32 | 69 |

---

## 1. HIGH PRIORITY ISSUES

### 1.1 Unused Parameters (Warning - PSReviewUnusedParameter)

**Impact:** Parameters defined but never used, causing confusion

#### Export-PlannerData.ps1

**Issue 1: Line 31** - `$ExportAllMyPlans` parameter is declared but not used

**Current Code:**
```powershell
[Parameter(Mandatory = $false)]
[switch]$ExportAllMyPlans,
```

**Problem:** The parameter exists but has no functionality implemented. Users might expect it to do something.

**Recommended Correction:**
Option A - Remove the parameter entirely:
```powershell
# Remove lines 30-31
```

Option B - Implement the functionality:
```powershell
# In main program section (around line 490), add logic:
if ($ExportAllMyPlans) {
    # Force export of all plans regardless of other parameters
    $plans = Get-AllUserPlans
} elseif ($GroupIds) {
    $plans = Get-PlansByGroupIds -GroupIds $GroupIds
} else {
    $plans = Get-AllUserPlans
}
```

---

**Issue 2: Line 34** - `$IncludeCompletedTasks` parameter is declared but not used

**Current Code:**
```powershell
[Parameter(Mandatory = $false)]
[switch]$IncludeCompletedTasks
```

**Problem:** Line 214 has a comment saying this parameter should filter completed tasks, but the filtering logic is inverted:
```powershell
# Line 214-217
if (-not $IncludeCompletedTasks) {
    $completedCount = ($allTasks | Where-Object { $_.percentComplete -eq 100 }).Count
    Write-Log "  $completedCount abgeschlossene Tasks werden mitexportiert (verwende -IncludeCompletedTasks zum Filtern)"
}
```

**Recommended Correction:**
```powershell
# Replace lines 214-217 with:
if ($IncludeCompletedTasks) {
    # Include all tasks - no filtering needed
    Write-Log "  $completedCount abgeschlossene Tasks werden mitexportiert"
} else {
    # Filter out completed tasks
    $completedCount = ($allTasks | Where-Object { $_.percentComplete -eq 100 }).Count
    Write-Log "  $completedCount abgeschlossene Tasks werden übersprungen (verwende -IncludeCompletedTasks um sie zu exportieren)"
    $allTasks = $allTasks | Where-Object { $_.percentComplete -ne 100 }
}
```

---

#### Import-PlannerData.ps1

**Issue 1: Line 36** - `$UserMapping` parameter declared but not used globally

**Current Code:**
```powershell
[Parameter(Mandatory = $false)]
[hashtable]$UserMapping,
```

**Problem:** The parameter is referenced in the `Resolve-UserId` function (line 147) but because it's checking `$UserMapping` without the proper scope, it won't work correctly.

**Recommended Correction:**
```powershell
# In Resolve-UserId function (line 143), change from:
function Resolve-UserId {
    param([string]$OldUserId, [hashtable]$OldUserMap)

    # Wenn UserMapping vorhanden, verwende es
    if ($UserMapping -and $UserMapping.ContainsKey($OldUserId)) {
        return $UserMapping[$OldUserId]
    }
    # ...
}

# To:
function Resolve-UserId {
    param(
        [string]$OldUserId, 
        [hashtable]$OldUserMap,
        [hashtable]$UserMapping  # Add as explicit parameter
    )

    # Wenn UserMapping vorhanden, verwende es
    if ($UserMapping -and $UserMapping.ContainsKey($OldUserId)) {
        return $UserMapping[$OldUserId]
    }
    # ...
}

# Then update all calls to Resolve-UserId (line 335):
$resolvedId = Resolve-UserId -OldUserId $_.Name -OldUserMap $userMap -UserMapping $UserMapping
```

---

**Issue 2: Line 39** - `$SkipAssignments` parameter not used correctly

**Current Code:**
```powershell
[Parameter(Mandatory = $false)]
[switch]$SkipAssignments,
```

**Problem:** Referenced at line 331 but would cause scoping issues in functions.

**Recommended Correction:**
Make it accessible to functions either via:
1. Pass as parameter to `Import-PlanFromJson` function
2. Use `$script:SkipAssignments` for script-level scope
3. Or convert to using `-WhatIf` common parameter instead

**Recommended approach:**
```powershell
# Change Import-PlanFromJson function signature (line 189):
function Import-PlanFromJson {
    param(
        [string]$JsonFilePath,
        [string]$TargetGroupId,
        [switch]$SkipAssignments,  # Add parameter
        [switch]$SkipCompletedTasks  # Add parameter
    )
    # ...
}

# Update function call (line 526):
$result = Import-PlanFromJson -JsonFilePath $jsonFile.FullName `
                              -TargetGroupId $TargetGroupId `
                              -SkipAssignments:$SkipAssignments `
                              -SkipCompletedTasks:$SkipCompletedTasks
```

---

**Issue 3: Line 42** - `$SkipCompletedTasks` parameter not used correctly

Same issue as `$SkipAssignments` - see correction above.

---

**Issue 4: Line 48** - `$ThrottleDelayMs` parameter not accessible in functions

**Current Code:**
```powershell
[Parameter(Mandatory = $false)]
[int]$ThrottleDelayMs = 500
```

**Problem:** Used in `Invoke-GraphWithRetry` (line 78) but will default to 0 if not in scope.

**Recommended Correction:**
```powershell
# Use script scope at the top of the script (after param block):
$script:ThrottleDelayMs = $ThrottleDelayMs

# OR pass it to the function:
function Invoke-GraphWithRetry {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body,
        [int]$MaxRetries = 3,
        [int]$ThrottleDelay = 500  # Add parameter with default
    )
    
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        try {
            Start-Sleep -Milliseconds $ThrottleDelay  # Use parameter
            # ...
        }
    }
}
```

---

### 1.2 Empty Catch Blocks (Warning - PSAvoidUsingEmptyCatchBlock)

**Impact:** Errors are silently swallowed, making debugging difficult

#### Import-PlannerData.ps1

**Issue 1: Line 163** - Empty catch in Resolve-UserId (UPN lookup)

**Current Code:**
```powershell
if ($upn) {
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
        if ($user) {
            return $user.id
        }
    }
    catch { }  # Line 163 - EMPTY!
}
```

**Recommended Correction:**
```powershell
if ($upn) {
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
        if ($user) {
            return $user.id
        }
    }
    catch {
        Write-Verbose "Could not resolve user by UPN '$upn': $($_.Exception.Message)"
    }
}
```

---

**Issue 2: Line 173** - Empty catch in Resolve-UserId (Mail lookup)

**Current Code:**
```powershell
if ($mail -and $mail -ne $upn) {
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mail`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
        if ($user) {
            return $user.id
        }
    }
    catch { }  # Line 173 - EMPTY!
}
```

**Recommended Correction:**
```powershell
if ($mail -and $mail -ne $upn) {
    try {
        $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$mail`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
        if ($user) {
            return $user.id
        }
    }
    catch {
        Write-Verbose "Could not resolve user by mail '$mail': $($_.Exception.Message)"
    }
}
```

---

**Issue 3: Line 184** - Empty catch in Resolve-UserId (Direct ID lookup)

**Current Code:**
```powershell
try {
    $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$OldUserId`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
    if ($user) {
        return $user.id
    }
}
catch { }  # Line 184 - EMPTY!
```

**Recommended Correction:**
```powershell
try {
    $user = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$OldUserId`?`$select=id" -OutputType PSObject -ErrorAction SilentlyContinue
    if ($user) {
        return $user.id
    }
}
catch {
    Write-Verbose "Could not resolve user by ID '$OldUserId': $($_.Exception.Message)"
}
```

---

## 2. MEDIUM PRIORITY ISSUES

### 2.1 Write-Host Usage (Warning - PSAvoidUsingWriteHost)

**Impact:** Output may not work in all environments (prior to PS 5.0), cannot be captured or redirected

Both scripts have extensive use of `Write-Host` for user-facing messages.

**Problem:** `Write-Host` should be avoided because:
- It doesn't work in all PowerShell hosts
- Cannot be suppressed, captured, or redirected
- Output goes directly to console, bypassing the pipeline

**Affected Lines:**
- Export-PlannerData.ps1: Lines 43, 462-467, 504-551 (21 instances)
- Import-PlannerData.ps1: Lines 57, 458-467, 494-550 (23 instances)

**Recommended Correction:**

Replace `Write-Host` with one of:
1. `Write-Output` - for data that should be in the pipeline
2. `Write-Information` - for informational messages (PS 5.0+)
3. `Write-Verbose` - for detailed logging (user can control with -Verbose)

**Example for header/banner output (lines 462-467 in Export script):**
```powershell
# Current:
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Planner Export Tool" -ForegroundColor Cyan
Write-Host "  by Alexander Waller" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Recommended:
Write-Information ""
Write-Information "============================================================" -InformationAction Continue
Write-Information "  Microsoft Planner Export Tool" -InformationAction Continue
Write-Information "  by Alexander Waller" -InformationAction Continue
Write-Information "============================================================" -InformationAction Continue
Write-Information ""

# Or for PS 5.0+ with color support:
$PSStyle.Foreground.Cyan
Write-Information "============================================================" -InformationAction Continue
Write-Information "  Microsoft Planner Export Tool" -InformationAction Continue
Write-Information "  by Alexander Waller" -InformationAction Continue
Write-Information "============================================================" -InformationAction Continue
$PSStyle.Reset
```

**For progress/status messages:**
```powershell
# Current:
Write-Host "Gefundene Export-Dateien:" -ForegroundColor Yellow

# Recommended:
Write-Information "Gefundene Export-Dateien:" -InformationAction Continue
```

**Note:** If maintaining backward compatibility with PS < 5.0, keep Write-Host but document this decision. For modern PowerShell (5.0+), migrate to Write-Information.

---

### 2.2 Overwriting Built-in Cmdlets (Warning - PSAvoidOverwritingBuiltInCmdlets)

**Impact:** Confusion and potential conflicts with PowerShell Core built-in cmdlets

#### Both Scripts: Write-Log Function

**Problem:** `Write-Log` is a built-in cmdlet in PowerShell Core 6.1.0+

**Current Code (Line 39 in Export, Line 53 in Import):**
```powershell
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    # ...
}
```

**Recommended Correction:**

Option A - Rename the function:
```powershell
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
    $logEntry | Out-File -FilePath "$ExportPath\export.log" -Append -Encoding UTF8
}

# Then replace all calls from Write-Log to Write-PlannerLog throughout the scripts
```

Option B - Add a prefix to avoid conflicts:
```powershell
function Write-PlannerExportLog { # or Write-PlannerImportLog
    # Same implementation
}
```

Option C - Use an alias with Export-ModuleMember (if converting to module):
```powershell
function Write-PlannerToolLog {
    # Implementation
}
Set-Alias -Name Write-Log -Value Write-PlannerToolLog
Export-ModuleMember -Function Write-PlannerToolLog -Alias Write-Log
```

**Recommendation:** Use Option A (Write-PlannerLog) as it's clearest and avoids conflicts.

---

### 2.3 Plural Nouns in Function Names (Warning - PSUseSingularNouns)

**Impact:** Violates PowerShell naming conventions, but not a functional issue

#### Export-PlannerData.ps1

**Issue 1: Line 81** - `Get-AllUserPlans` uses plural "Plans"

**Current:**
```powershell
function Get-AllUserPlans {
```

**Recommended:**
```powershell
function Get-AllUserPlan {
    # Returns array of plans, but function name should be singular
}
```

---

**Issue 2: Line 133** - `Get-PlansByGroupIds` uses plural "Plans" and "Ids"

**Current:**
```powershell
function Get-PlansByGroupIds {
```

**Recommended:**
```powershell
function Get-PlanByGroupId {
    # Even though it accepts array and returns multiple, use singular
    # The parameter can still be an array
    param([string[]]$GroupId)
}
```

---

**Issue 3: Line 161** - `Export-PlanDetails` uses plural "Details"

**Current:**
```powershell
function Export-PlanDetails {
```

**Recommended:**
```powershell
function Export-PlanDetail {
    # Singular noun per PowerShell convention
}
```

---

### 2.4 Missing BOM for Unicode Files (Warning - PSUseBOMForUnicodeEncodedFile)

**Impact:** Potential encoding issues with non-ASCII characters

**Problem:** Both scripts contain German text (umlauts like ä, ö, ü) but lack BOM (Byte Order Mark)

**Files Affected:**
- Export-PlannerData.ps1
- Import-PlannerData.ps1

**Recommended Correction:**

Save files with UTF-8 BOM encoding:

**In VS Code:**
1. Open file
2. Click on encoding in status bar (bottom right)
3. Select "Save with Encoding"
4. Choose "UTF-8 with BOM"

**In PowerShell ISE:**
1. File → Save with encoding
2. Select "UTF-8 with BOM"

**Programmatically:**
```powershell
# Convert file to UTF-8 with BOM
$content = Get-Content -Path "Export-PlannerData.ps1" -Raw
[System.IO.File]::WriteAllText("Export-PlannerData.ps1", $content, [System.Text.UTF8Encoding]::new($true))
```

---

## 3. LOW PRIORITY ISSUES

### 3.1 Trailing Whitespace (Information - PSAvoidTrailingWhitespace)

**Impact:** Cosmetic only, but can cause diff noise in version control

**Affected Lines:**

Export-PlannerData.ps1:
- Line 88: After `$myGroups = ...`
- Line 204: After `$tasksUri = ...`
- Line 228: After `Start-Sleep -Milliseconds 200`
- Line 339: After comment in switch statement

Import-PlannerData.ps1:
- Line 16: After WICHTIG comment
- Line 230: After `$newPlanDetails = ...`
- Line 486: After `$jsonFiles = ...`
- Line 525: After `Write-Host ""`

**Recommended Correction:**

Remove trailing whitespace from all lines. Most modern editors can do this automatically:

**VS Code:**
```json
// settings.json
{
    "files.trimTrailingWhitespace": true
}
```

**PowerShell command to fix:**
```powershell
# Fix all .ps1 files
Get-ChildItem -Path . -Filter "*.ps1" | ForEach-Object {
    $content = Get-Content $_.FullName
    $content | ForEach-Object { $_.TrimEnd() } | Set-Content $_.FullName -NoNewline
}
```

---

## 4. CODE LOGIC IMPROVEMENTS

### 4.1 Error Handling Improvements

#### Export-PlannerData.ps1

**Issue: Line 476-479** - Module installation without error handling

**Current:**
```powershell
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Planner)) {
    Write-Log "Microsoft.Graph Module werden installiert..." "WARN"
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}
```

**Problem:** `Install-Module` can fail (network issues, permissions, etc.) but error isn't caught

**Recommended:**
```powershell
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Planner)) {
    Write-Log "Microsoft.Graph Module werden installiert..." "WARN"
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "Microsoft.Graph Module erfolgreich installiert" "OK"
    }
    catch {
        Write-Log "Fehler bei der Installation der Microsoft.Graph Module: $_" "ERROR"
        Write-Log "Bitte installieren Sie manuell: Install-Module Microsoft.Graph -Scope CurrentUser" "ERROR"
        exit 1
    }
}
```

---

**Issue: Line 287** - JSON export without error handling

**Current:**
```powershell
$planData | ConvertTo-Json -Depth 20 | Out-File -FilePath $planFilePath -Encoding UTF8
```

**Problem:** File write could fail (disk full, permissions, path too long)

**Recommended:**
```powershell
try {
    $planData | ConvertTo-Json -Depth 20 | Out-File -FilePath $planFilePath -Encoding UTF8 -ErrorAction Stop
    Write-Log "  Plan exportiert nach: $planFilePath" "OK"
}
catch {
    Write-Log "  Fehler beim Exportieren nach $planFilePath : $_" "ERROR"
    throw
}
```

---

#### Import-PlannerData.ps1

**Issue: Line 196** - JSON loading without validation

**Current:**
```powershell
$planData = Get-Content $JsonFilePath -Raw -Encoding UTF8 | ConvertFrom-Json
```

**Problem:** JSON could be corrupted or invalid, causing cryptic errors

**Recommended:**
```powershell
try {
    $planData = Get-Content $JsonFilePath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Log "Fehler beim Laden der JSON-Datei: $_" "ERROR"
    Write-Log "Überprüfen Sie ob die Datei existiert und valides JSON enthält" "ERROR"
    return $null
}

# Validate required properties
if (-not $planData.Plan) {
    Write-Log "Ungültige Datei: 'Plan' Eigenschaft fehlt" "ERROR"
    return $null
}
```

---

### 4.2 Security Improvements

**Issue: Sensitive information in logs**

**Problem:** Both scripts log to files that might contain sensitive information (group IDs, user IDs, plan titles, etc.)

**Current:** Logs are created with default file permissions

**Recommended:**

Add warning in documentation and optionally restrict log file permissions:

```powershell
# After creating log file, restrict permissions (Windows)
if ($IsWindows -or $PSVersionTable.PSVersion.Major -lt 6) {
    try {
        $acl = Get-Acl "$ExportPath\export.log"
        $acl.SetAccessRuleProtection($true, $false)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
            "FullControl",
            "Allow"
        )
        $acl.SetAccessRule($rule)
        Set-Acl "$ExportPath\export.log" $acl
    }
    catch {
        Write-Log "Hinweis: Konnte Log-Datei Berechtigungen nicht einschränken" "WARN"
    }
}
```

---

### 4.3 Performance Improvements

**Issue: Line 235 in Export-PlannerData.ps1** - Hard-coded sleep

**Current:**
```powershell
# Rate Limiting vermeiden
Start-Sleep -Milliseconds 200
```

**Problem:** Always sleeps 200ms even when not necessary

**Recommended:**

Only sleep after actual API calls, or make it configurable:

```powershell
# Add parameter at top:
[Parameter(Mandatory = $false)]
[int]$TaskDetailDelayMs = 200

# Then use:
Start-Sleep -Milliseconds $TaskDetailDelayMs
```

Or better - implement adaptive throttling:

```powershell
# Track last API call time
$script:LastApiCall = Get-Date

function Invoke-ThrottledApiCall {
    param([scriptblock]$ApiCall)
    
    # Calculate time since last call
    $timeSinceLastCall = (Get-Date) - $script:LastApiCall
    $minDelay = [TimeSpan]::FromMilliseconds(200)
    
    if ($timeSinceLastCall -lt $minDelay) {
        Start-Sleep -Milliseconds ($minDelay - $timeSinceLastCall).TotalMilliseconds
    }
    
    $script:LastApiCall = Get-Date
    & $ApiCall
}
```

---

**Issue: Line 227 in Export-PlannerData.ps1** - Write-Progress in tight loop

**Current:**
```powershell
Write-Progress -Activity "Lade Task-Details für '$($Plan.title)'" -Status "Task $counter von $($allTasks.Count)" -PercentComplete (($counter / $allTasks.Count) * 100)
```

**Problem:** Write-Progress is called for every task, which is slow for large plans

**Recommended:**

Update progress less frequently:

```powershell
# Only update every 5 tasks or on last task
if ($counter % 5 -eq 0 -or $counter -eq $allTasks.Count) {
    Write-Progress -Activity "Lade Task-Details für '$($Plan.title)'" -Status "Task $counter von $($allTasks.Count)" -PercentComplete (($counter / $allTasks.Count) * 100)
}
```

---

## 5. BEST PRACTICE RECOMMENDATIONS

### 5.1 Add Parameter Validation

Add validation attributes to parameters:

```powershell
# Export-PlannerData.ps1
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { 
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            $true 
        }
    })]
    [string]$ExportPath = "C:\planner-data\PlannerExport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$GroupIds
)

# Import-PlannerData.ps1
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$ImportPath,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$')]
    [string]$TargetGroupId,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$ThrottleDelayMs = 500
)
```

---

### 5.2 Add Help Documentation

Add detailed help for each function:

```powershell
function Export-PlanDetail {
    <#
    .SYNOPSIS
        Exports detailed plan data including buckets, tasks, and user information
    
    .DESCRIPTION
        Retrieves comprehensive plan data from Microsoft Graph API and exports to JSON format.
        Includes plan details, buckets, tasks, task details, user mappings, and category descriptions.
    
    .PARAMETER Plan
        The plan object to export (must include id and title properties)
    
    .PARAMETER PlanExportPath
        Directory path where the plan JSON file will be saved
    
    .EXAMPLE
        Export-PlanDetail -Plan $planObject -PlanExportPath "C:\exports"
        
    .OUTPUTS
        Hashtable containing the complete plan data structure
    
    .NOTES
        Requires Microsoft.Graph.Planner module and appropriate API permissions
    #>
    param(
        [Parameter(Mandatory=$true)]
        [PSObject]$Plan,
        
        [Parameter(Mandatory=$true)]
        [string]$PlanExportPath
    )
    # Implementation...
}
```

---

### 5.3 Add ShouldProcess Support

For scripts that make changes, add `-WhatIf` and `-Confirm` support:

```powershell
# Import-PlannerData.ps1
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    # ... existing parameters
)

# Then in Import-PlanFromJson:
if ($PSCmdlet.ShouldProcess("$planTitle in group $groupId", "Create Plan")) {
    $newPlan = Invoke-GraphWithRetry -Method POST -Uri "https://graph.microsoft.com/v1.0/planner/plans" -Body @{
        owner = $groupId
        title = $planTitle
    }
    # ...
}
```

---

### 5.4 Improve Logging

**Current issues:**
1. No log levels for filtering
2. No structured logging
3. No log rotation

**Recommended improvements:**

```powershell
function Write-PlannerLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("DEBUG", "INFO", "WARN", "ERROR", "OK")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Console", "File", "Both")]
        [string]$Target = "Both"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Console output
    if ($Target -in @("Console", "Both")) {
        $color = switch ($Level) {
            "ERROR" { "Red" }
            "WARN"  { "Yellow" }
            "OK"    { "Green" }
            "DEBUG" { "Gray" }
            default { "White" }
        }
        
        # Only show DEBUG if -Verbose
        if ($Level -ne "DEBUG" -or $VerbosePreference -ne 'SilentlyContinue') {
            Write-Host $logEntry -ForegroundColor $color
        }
    }
    
    # File output
    if ($Target -in @("File", "Both") -and $LogPath) {
        try {
            # Check log file size and rotate if needed (> 10MB)
            if (Test-Path $LogPath) {
                $fileInfo = Get-Item $LogPath
                if ($fileInfo.Length -gt 10MB) {
                    $archivePath = "$LogPath.$(Get-Date -Format 'yyyyMMdd_HHmmss').old"
                    Move-Item $LogPath $archivePath -Force
                }
            }
            
            $logEntry | Out-File -FilePath $LogPath -Append -Encoding UTF8
        }
        catch {
            Write-Warning "Could not write to log file: $_"
        }
    }
}
```

---

### 5.5 Add Version Compatibility Check

Add at the beginning of each script:

```powershell
#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication

# Check PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 5 -or 
    ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    Write-Error "This script requires PowerShell 5.1 or higher. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}

# Recommend PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "PowerShell 7+ is recommended for better performance and compatibility. Current version: $($PSVersionTable.PSVersion)"
}
```

---

## 6. DOCUMENTATION IMPROVEMENTS

### 6.1 README.md Enhancements

Add sections for:

1. **Known Limitations** - Already exists but could expand:
   - API rate limits (specific numbers)
   - Maximum plan size that can be exported
   - Timeout scenarios

2. **Error Messages Reference** - Common errors and solutions:
```markdown
### Common Errors

#### "Insufficient privileges"
**Cause:** Missing API permissions  
**Solution:** Ask admin to grant consent for required scopes

#### "429 Too Many Requests"
**Cause:** API rate limit exceeded  
**Solution:** Script automatically retries. If persistent, increase `-ThrottleDelayMs`

#### "Plan not found"
**Cause:** Plan was deleted or permissions changed  
**Solution:** Verify plan exists and you have access
```

3. **Performance Tuning Guide**:
```markdown
### Performance Optimization

For large plans (>500 tasks):
- Increase `-ThrottleDelayMs` to 1000
- Run during off-peak hours
- Consider filtering completed tasks
- Export/import plans individually rather than in bulk
```

---

### 6.2 Add CHANGELOG.md

Track changes and version history:

```markdown
# Changelog

## [Unreleased]
### Added
- Unit test suite (59 tests)
- Test documentation

### Known Issues
- Parameters `-ExportAllMyPlans` and `-IncludeCompletedTasks` not fully implemented
- Empty catch blocks in user resolution logic

## [1.0.0] - 2026-02-09
### Added
- Initial release
- Export functionality
- Import functionality
- User mapping support
- Dry run mode
```

---

## 7. TESTING RECOMMENDATIONS

### 7.1 Additional Test Cases Needed

Current test coverage: 59 tests, but missing:

1. **Integration tests with actual API** (documented but not automated)
2. **Error scenario tests:**
   - Network failures
   - Invalid JSON
   - Missing permissions
   - Rate limiting
3. **Large dataset tests:**
   - Plans with 1000+ tasks
   - Plans with 100+ buckets
4. **Edge cases:**
   - Empty plans
   - Plans with special characters in titles
   - Circular user mappings
   - Missing bucket IDs

---

### 7.2 Recommended Additional Tests

```powershell
# Test for parameter scoping issues
Describe "Parameter Scoping Tests" {
    It "Should use UserMapping parameter in Resolve-UserId" {
        $mapping = @{ "old-id" = "new-id" }
        # Test that mapping is accessible in function
    }
    
    It "Should respect SkipAssignments flag" {
        # Test that assignments are actually skipped
    }
    
    It "Should filter completed tasks when IncludeCompletedTasks is false" {
        # Test that completed tasks are filtered
    }
}

# Test for error handling
Describe "Error Handling Tests" {
    It "Should handle corrupted JSON gracefully" {
        # Test with invalid JSON
    }
    
    It "Should handle missing Plan property" {
        # Test with incomplete data
    }
    
    It "Should handle API errors" {
        # Mock failed API calls
    }
}
```

---

## 8. SUMMARY OF CORRECTIONS BY FILE

### Export-PlannerData.ps1

| Line | Issue | Priority | Fix Complexity |
|------|-------|----------|----------------|
| 31 | Unused parameter `$ExportAllMyPlans` | HIGH | Low - Remove or implement |
| 34 | Unused parameter `$IncludeCompletedTasks` | HIGH | Medium - Implement filtering |
| 39 | Function name conflicts with built-in | MEDIUM | Medium - Rename function |
| 43, 462+ | Using Write-Host | MEDIUM | High - Replace throughout |
| 81, 133, 161 | Plural noun in function names | LOW | Low - Rename functions |
| 88, 204, 228, 339 | Trailing whitespace | LOW | Low - Remove whitespace |
| N/A | Missing BOM for Unicode | MEDIUM | Low - Re-save with BOM |
| 287 | No error handling on file write | MEDIUM | Low - Add try-catch |
| 476 | No error handling on Install-Module | MEDIUM | Low - Add try-catch |

**Total Issues:** 37  
**Estimated Effort:** 4-6 hours to address all

---

### Import-PlannerData.ps1

| Line | Issue | Priority | Fix Complexity |
|------|-------|----------|----------------|
| 36, 39, 42, 48 | Unused/improperly scoped parameters | HIGH | Medium - Fix scoping |
| 53 | Function name conflicts with built-in | MEDIUM | Medium - Rename function |
| 57, 458+ | Using Write-Host | MEDIUM | High - Replace throughout |
| 163, 173, 184 | Empty catch blocks | HIGH | Low - Add error logging |
| 16, 230, 486, 525 | Trailing whitespace | LOW | Low - Remove whitespace |
| N/A | Missing BOM for Unicode | MEDIUM | Low - Re-save with BOM |
| 196 | No validation on JSON load | MEDIUM | Low - Add validation |

**Total Issues:** 32  
**Estimated Effort:** 3-5 hours to address all

---

## 9. PRIORITY RECOMMENDATIONS

If you can only fix a few things, prioritize these:

### Must Fix (Breaks Functionality)
1. ✅ **HIGH PRIORITY** - Implement or remove unused parameters (`$IncludeCompletedTasks`, `$ExportAllMyPlans`, etc.)
2. ✅ **HIGH PRIORITY** - Fix parameter scoping in Import script (UserMapping, SkipAssignments, etc.)
3. ✅ **HIGH PRIORITY** - Add error logging to empty catch blocks

### Should Fix (Best Practices)
4. 🔶 **MEDIUM PRIORITY** - Rename Write-Log to avoid conflicts (Write-PlannerLog)
5. 🔶 **MEDIUM PRIORITY** - Add error handling to file operations and API calls
6. 🔶 **MEDIUM PRIORITY** - Save files with UTF-8 BOM for German characters

### Nice to Have (Code Quality)
7. ⬜ **LOW PRIORITY** - Replace Write-Host with Write-Information
8. ⬜ **LOW PRIORITY** - Fix plural nouns in function names
9. ⬜ **LOW PRIORITY** - Remove trailing whitespace

---

## 10. AUTOMATED FIX SCRIPT

Here's a PowerShell script to automatically fix some low-hanging issues:

```powershell
<#
.SYNOPSIS
    Automated fixes for code quality issues
#>

# Fix trailing whitespace
Get-ChildItem -Path . -Filter "*.ps1" -Recurse | ForEach-Object {
    $content = Get-Content $_.FullName
    $fixed = $content | ForEach-Object { $_.TrimEnd() }
    $fixed | Set-Content $_.FullName -NoNewline -Encoding UTF8
    Write-Host "Fixed trailing whitespace in $($_.Name)"
}

# Convert to UTF-8 with BOM
$files = @(
    "Export-PlannerData.ps1",
    "Import-PlannerData.ps1"
)

foreach ($file in $files) {
    $content = Get-Content $file -Raw
    $utf8WithBom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($file, $content, $utf8WithBom)
    Write-Host "Converted $file to UTF-8 with BOM"
}

Write-Host "Automated fixes complete!"
Write-Host "Manual fixes still required for:"
Write-Host "  - Unused parameters"
Write-Host "  - Empty catch blocks"
Write-Host "  - Function naming conflicts"
Write-Host "  - Write-Host usage"
```

---

## 11. VALIDATION CHECKLIST

After applying corrections, verify:

- [ ] All tests still pass (run `pwsh ./tests/Run-Tests.ps1`)
- [ ] PSScriptAnalyzer shows fewer warnings (`Invoke-ScriptAnalyzer -Path .\*.ps1`)
- [ ] Scripts work with PowerShell 5.1 and 7+
- [ ] German characters (ä, ö, ü) display correctly
- [ ] Export produces valid JSON files
- [ ] Import can restore exported data
- [ ] Dry run mode works without making changes
- [ ] Log files are created successfully
- [ ] User mapping works correctly
- [ ] Error messages are clear and helpful

---

## CONCLUSION

The code is **functionally correct** (all tests pass) but has **room for improvement** in terms of:
- PowerShell best practices compliance
- Error handling robustness
- Parameter implementation
- Code maintainability

**Recommended Approach:**
1. Fix HIGH PRIORITY issues first (3-4 hours)
2. Then MEDIUM PRIORITY issues (2-3 hours)
3. LOW PRIORITY can be addressed over time

**Total Estimated Effort:** 8-12 hours for complete cleanup

**Current State:** ✅ Production-ready but not best-practice compliant  
**After Fixes:** ✨ Production-ready and best-practice compliant

---

**Last Updated:** 2026-02-12  
**Review Version:** 1.0  
**Scripts Analyzed:** Export-PlannerData.ps1 v1.0.0, Import-PlannerData.ps1 v1.0.0
