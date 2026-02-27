<#
.SYNOPSIS
    Prueft ob die erforderlichen Microsoft Graph Berechtigungen vorhanden sind.

.DESCRIPTION
    Dieses Skript verbindet sich mit Microsoft Graph und prueft,
    ob die fuer "Microsoft Graph Command Line Tools" benoetigten
    Berechtigungen (Scopes) vorhanden sind.

.NOTES
    Autor:  Alexander Waller
    Datum:  2026-02-27
    Benoetigte Module: Microsoft.Graph.Authentication
#>

#Requires -Version 5.1

# ISE-Erkennung: WAM deaktivieren (ISE unterstuetzt WAM nicht zuverlaessig)
if ($host.Name -eq 'Windows PowerShell ISE Host') {
    Write-Host "[!] PowerShell ISE erkannt - WAM wird deaktiviert." -ForegroundColor Yellow
    Write-Host "    Empfehlung: pwsh.exe oder powershell.exe verwenden." -ForegroundColor Yellow
    Write-Host ""
    $env:MSAL_ENABLE_WAM = "0"
}

# --- Konfiguration: Benoetigte Berechtigungen ---
$RequiredPermissions = @(
    "User.ReadBasic.All"        # Read all users' basic profiles
    "Group.Read.All"            # Read all groups
    "offline_access"            # Maintain access to data you have given it access to
    "Tasks.Read"                # Read user's tasks and task lists
    "User.Read"                 # Sign in and read user profile
    "Tasks.ReadWrite"           # Create, read, update, and delete user's tasks and task lists
)

# --- Farb-Hilfsfunktionen ---
function Write-Status {
    param(
        [string]$Permission,
        [bool]$Granted
    )
    if ($Granted) {
        Write-Host "  [OK] " -ForegroundColor Green -NoNewline
        Write-Host $Permission
    } else {
        Write-Host "  [FEHLT] " -ForegroundColor Red -NoNewline
        Write-Host $Permission
    }
}

# --- Hauptlogik ---
Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Graph Berechtigungs-Pruefung" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# Pruefen ob Microsoft.Graph Modul installiert ist
$graphModule = Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication"
if (-not $graphModule) {
    Write-Host "[!] Microsoft.Graph.Authentication Modul nicht gefunden." -ForegroundColor Yellow
    Write-Host "    Installieren mit: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host ""
    
    $install = Read-Host "Modul jetzt installieren? (j/n)"
    if ($install -eq 'j') {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    } else {
        Write-Host "Abbruch." -ForegroundColor Red
        throw "Microsoft.Graph.Authentication Modul nicht installiert."
    }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Pruefen ob bereits verbunden
$context = Get-MgContext

if (-not $context) {
    Write-Host "Verbinde mit Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "(Es wird ein Browser-Fenster geoeffnet)" -ForegroundColor Gray
    Write-Host ""
    
    try {
        Connect-MgGraph -Scopes $RequiredPermissions -NoWelcome -ErrorAction Stop
        $context = Get-MgContext
    }
    catch {
        Write-Host "Browser-Authentifizierung fehlgeschlagen, verwende Device Code Flow..." -ForegroundColor Yellow
        try {
            Connect-MgGraph -Scopes $RequiredPermissions -UseDeviceCode -NoWelcome -ErrorAction Stop
            $context = Get-MgContext
        }
        catch {
            Write-Host "[FEHLER] Verbindung fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
            throw "Authentifizierung fehlgeschlagen: $($_.Exception.Message)"
        }
    }
}

if (-not $context) {
    Write-Host "[FEHLER] Keine aktive Graph-Verbindung." -ForegroundColor Red
    throw "Keine aktive Graph-Verbindung."
}

# --- Verbindungsinfo anzeigen ---
Write-Host "Verbunden als:" -ForegroundColor Green
Write-Host "  Konto:    $($context.Account)" -ForegroundColor White
Write-Host "  Tenant:   $($context.TenantId)" -ForegroundColor White
Write-Host "  App:      $($context.AppName)" -ForegroundColor White
Write-Host "  Auth-Typ: $($context.AuthType)" -ForegroundColor White
Write-Host ""

# --- Aktuelle Scopes auslesen ---
$currentScopes = $context.Scopes

Write-Host "Aktuelle Scopes ($($currentScopes.Count) vorhanden):" -ForegroundColor Cyan
Write-Host "--------------------------------------------------------"

# --- Berechtigungen pruefen ---
$missing = @()
$granted = @()

foreach ($perm in $RequiredPermissions) {
    # Case-insensitive Vergleich
    $found = $currentScopes | Where-Object { $_ -ieq $perm }
    if ($found) {
        $granted += $perm
        Write-Status -Permission $perm -Granted $true
    } else {
        $missing += $perm
        Write-Status -Permission $perm -Granted $false
    }
}

# Zusaetzliche Scopes anzeigen (die nicht in der Required-Liste sind)
$extraScopes = $currentScopes | Where-Object { 
    $scope = $_
    -not ($RequiredPermissions | Where-Object { $_ -ieq $scope })
}

if ($extraScopes.Count -gt 0) {
    Write-Host ""
    Write-Host "Zusaetzliche vorhandene Scopes:" -ForegroundColor Gray
    foreach ($scope in $extraScopes) {
        Write-Host "  [+] $scope" -ForegroundColor DarkGray
    }
}

# --- Zusammenfassung ---
Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  Zusammenfassung" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Erforderlich:  $($RequiredPermissions.Count)" -ForegroundColor White
Write-Host "  Vorhanden:     $($granted.Count)" -ForegroundColor Green
Write-Host "  Fehlend:       $($missing.Count)" -ForegroundColor $(if ($missing.Count -gt 0) { "Red" } else { "Green" })

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "ACHTUNG: Fehlende Berechtigungen erfordern Admin-Genehmigung!" -ForegroundColor Yellow
    Write-Host "Fehlende Scopes:" -ForegroundColor Yellow
    foreach ($m in $missing) {
        Write-Host "  - $m" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Loesung: Admin-Consent anfordern oder mit allen Scopes neu verbinden:" -ForegroundColor Yellow
    Write-Host "  Disconnect-MgGraph" -ForegroundColor Gray
    Write-Host "  Connect-MgGraph -Scopes '$($RequiredPermissions -join "','")'" -ForegroundColor Gray
} else {
    Write-Host ""
    Write-Host "Alle erforderlichen Berechtigungen sind vorhanden!" -ForegroundColor Green
}

Write-Host ""