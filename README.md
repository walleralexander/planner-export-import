# Microsoft Planner Export/Import Tool

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-orange)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)
![Tests](https://img.shields.io/badge/Tests-93%2F97%20passing-green)
![Code Quality](https://img.shields.io/badge/Code%20Quality-8.2%2F10-yellowgreen)

---

> ⚠️ **WICHTIG:** Bitte lesen Sie die [WARNING.md](WARNING.md) für wichtige Hinweise zu Haftungsausschluss und Testumgebung!
> 📋 **CODE REVIEW:** Siehe [CODE_REVIEW_SUMMARY.md](CODE_REVIEW_SUMMARY.md) für Code-Qualitätsanalyse und Verbesserungsvorschläge

---

## 🆕 Neu in Version 1.1.0 (Februar 2026)

### Flexible Export-Modi

Das Export-Skript unterstützt jetzt **zwei verschiedene Modi**:

1. **👤 User-basiert** (`-UseCurrentUser`): Exportiert alle Pläne des aktuell angemeldeten Benutzers

   ```powershell
   .\Export-PlannerData.ps1 -UseCurrentUser
   ```

2. **🏢 Gruppen-basiert**: Exportiert Pläne aus spezifischen M365-Gruppen/SharePoint-Seiten

   - **Nach Gruppennamen** (`-GroupNames`): Sucht Gruppen nach ihrem Display-Namen

     ```powershell
     .\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha", "Marketing"
     ```

   - **Nach Gruppen-IDs** (`-GroupIds`): Direkte Angabe von Gruppen-IDs

     ```powershell
     .\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."
     ```

   - **Interaktiv** (`-Interactive`): Zeigt alle verfügbaren Gruppen zur Auswahl an

     ```powershell
     .\Export-PlannerData.ps1 -Interactive
     ```

### Weitere Verbesserungen

- ✅ Korrekte Umlaut-Darstellung in allen Ausgaben
- ✅ Verbesserte Fehlerprüfung und Null-Validierung
- ✅ Detailliertere Log-Meldungen mit Statusangaben
- ✅ Bessere Behandlung von fehlenden oder ungültigen Gruppen

---

## Übersicht

Dieses Toolset ermöglicht den fast vollständigen Export und Import von Microsoft Planner-Daten
über die Microsoft Graph API. Es wurde speziell für den Lizenzwechsel erstellt, um
Datenverluste zu vermeiden.

### Was wird exportiert/importiert?

| Datenpunkt | Export | Import |
| --- | --- | --- |
| Pläne | ✅ | ✅ |
| Buckets (Spalten) | ✅ | ✅ |
| Tasks (Aufgaben) | ✅ | ✅ |
| Beschreibungen | ✅ | ✅ |
| Checklisten | ✅ | ✅ |
| Zuweisungen (Personen) | ✅ | ✅ * |
| Priorität | ✅ | ✅ |
| Fälligkeitsdatum | ✅ | ✅ |
| Startdatum | ✅ | ✅ |
| Labels/Kategorien | ✅ | ✅ |
| Links/Referenzen | ✅ | ✅ |
| Fortschritt (%) | ✅ | ✅ |
| Kommentare | ❌ ** | ❌ |
| Dateianhänge | ✅ (als Link) | ✅ (als Link) |

\* Zuweisungen funktionieren nur wenn die Benutzer im neuen Tenant existieren (gleiche UPN/Mail)  
\** Kommentare sind über die Planner API nicht zugänglich (werden in Exchange gespeichert)

---

## Voraussetzungen

```powershell
winget install Microsoft.PowerShell
```

### 2. Microsoft Graph PowerShell Module

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### 3. Berechtigungen

Beim ersten Ausführen wird ein Browser-Fenster für die Anmeldung geöffnet.
Benötigte Berechtigungen (Delegated):

- `Group.Read.All` (Export) / `Group.ReadWrite.All` (Import)
- `Tasks.Read` (Export) / `Tasks.ReadWrite` (Import)
- `User.Read`
- `User.ReadBasic.All`

---

## Verwendung

### Export

#### 👤 User-basierte Export (Alle eigenen Pläne)

```powershell
# Alle Pläne des aktuellen Benutzers exportieren
.\Export-PlannerData.ps1 -UseCurrentUser

# Mit abgeschlossenen Tasks
.\Export-PlannerData.ps1 -UseCurrentUser -IncludeCompletedTasks
```

#### 🏢 Gruppen-basierte Export

**Nach Gruppennamen:**

```powershell
# Eine Gruppe
.\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha"

# Mehrere Gruppen
.\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha", "Marketing Team"
```

**Nach Gruppen-IDs:**

```powershell
.\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."
```

> **Tipp:** Die Gruppen-ID findet man in der URL wenn man die Gruppe in Outlook/Teams öffnet,
> oder über das Azure AD Portal unter Gruppen.

**Interaktive Auswahl:**

```powershell
.\Export-PlannerData.ps1 -Interactive
```

> Zeigt eine Liste aller verfügbaren M365-Gruppen zur Auswahl an.

#### Export in bestimmtes Verzeichnis

```powershell
.\Export-PlannerData.ps1 -UseCurrentUser -ExportPath "C:\Backup\Planner"
```

> Standardmäßig wird nach `C:\planner-data\PlannerExport_YYYYMMDD_HHMMSS` exportiert.

### Import

#### Alle exportierten Pläne importieren (gleiche Gruppen)

```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000"
```

#### In eine bestimmte Gruppe importieren

```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -TargetGroupId "neue-gruppe-id"
```

#### Probelauf (Dry Run) - zeigt was gemacht würde

```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -DryRun
```

#### Ohne Zuweisungen importieren

```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipAssignments
```

#### Abgeschlossene Tasks beim Import überspringen

```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipCompletedTasks
```

#### Benutzer-Mapping (wenn User-IDs sich ändern)

```powershell
$mapping = @{
    "alte-user-id-1" = "neue-user-id-1"
    "alte-user-id-2" = "neue-user-id-2"
}
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -UserMapping $mapping
```

---

## Exportierte Dateien

Nach dem Export enthält das Verzeichnis:

```Text
PlannerExport_20260209_143000/
├── _ExportIndex.json                    # Gesamtübersicht aller exportierten Pläne
├── export.log                           # Log-Datei des Exports
├── Projektplanung_2026.json             # Strukturierte Daten (für Import)
├── Projektplanung_2026_Zusammenfassung.txt  # Lesbare Textübersicht
├── IT_Infrastruktur.json
├── IT_Infrastruktur_Zusammenfassung.txt
└── ...
```

- **JSON-Dateien**: Enthalten alle strukturierten Daten für den Re-Import
- **Zusammenfassungs-Dateien**: Menschenlesbare Übersicht aller Tasks, Buckets etc.
- **ExportIndex**: Metadaten über den gesamten Export

---

## Empfohlene Vorgehensweise für den Lizenzwechsel

1. **VOR dem Wechsel:**

   ```powershell
   # Alle Pläne des aktuellen Benutzers exportieren
   .\Export-PlannerData.ps1 -UseCurrentUser -ExportPath "C:\Backup\Planner_PreMigration"

   # ODER: Spezifische Gruppen exportieren
   .\Export-PlannerData.ps1 -GroupNames "Projektteam", "Marketing" -ExportPath "C:\Backup\Planner_PreMigration"

   # Zusammenfassungen prüfen - stimmt alles?
   Get-ChildItem "C:\Backup\Planner_PreMigration\*Zusammenfassung*"
   ```

2. **Backup sichern:**
   - Export-Verzeichnis auf Netzlaufwerk oder externen Speicher kopieren
   - Sicherheitshalber ZIP erstellen

3. **NACH dem Wechsel:**

   ```powershell
   # Erst einen Probelauf:
   .\Import-PlannerData.ps1 -ImportPath "C:\Backup\Planner_PreMigration" -DryRun
   
   # Wenn alles OK, Import starten:
   .\Import-PlannerData.ps1 -ImportPath "C:\Backup\Planner_PreMigration"
   ```

4. **Stichproben prüfen:**
   - Öffne die importierten Pläne in Planner
   - Prüfe Buckets, Tasks, Beschreibungen, Checklisten
   - Prüfe Zuweisungen und Fälligkeitsdaten

---

## Fehlerbehebung

| Problem | Lösung |
| --- | --- |
| "Insufficient privileges" | Azure AD Admin muss die App-Berechtigungen freigeben |
| "429 Too Many Requests" | Script wartet automatisch, ggf. `-ThrottleDelayMs 1000` erhöhen |
| Zuweisungen fehlen | Benutzer existieren nicht im Tenant → `-SkipAssignments` oder UserMapping |
| Leerer Export | Prüfe ob der Account Planner-Lizenz hat und Mitglied der Gruppen ist |
| Kommentare fehlen | Kommentare sind über die API nicht exportierbar (Exchange-basiert) |
| "Unexpected token" / Parse-Fehler beim Start | Zeilenenden-Problem (LF statt CRLF) – siehe unten |

### Parse-Fehler: "Unexpected token" oder "Missing argument"

Wenn das Script beim Start sofort mit Parse-Fehlern abbricht, obwohl die Datei korrekt aussieht:

```text
Unexpected token 'User-basiert:' in expression or statement.
Missing argument in parameter list.
```

**Ursache:** PowerShell 5.1 auf Windows erwartet CRLF-Zeilenenden (`\r\n`). Wird die Datei
mit LF-only (`\n`, typisch bei Downloads über Linux/macOS oder bestimmte Browser) gespeichert,
erkennt PowerShell 5.1 den Block-Kommentar `<# ... #>` nicht korrekt und versucht,
den Kommentarinhalt als Code zu parsen.

**Lösung:** Datei neu von GitHub herunterladen. Das Repository enthält eine
[`.gitattributes`](.gitattributes)-Datei, die CRLF-Zeilenenden für alle `.ps1`-Dateien
erzwingt – ein frischer Download/Clone liefert automatisch die richtige Formatierung:

```powershell
# Option 1: Neu von GitHub klonen
git clone https://github.com/walleralexander/planner-export-import.git

# Option 2: Bestehenden Clone aktualisieren und Zeilenenden neu normalisieren
git pull
git rm --cached -r .
git reset --hard HEAD
```

Alternativ lässt sich die Korrektur auch direkt in PowerShell durchführen:

```powershell
# Zeilenenden in einer einzelnen Datei auf CRLF setzen
$file = "Export-PlannerData.ps1"
$content = [System.IO.File]::ReadAllText($file)
$content = $content.Replace("`r`n", "`n").Replace("`n", "`r`n")
[System.IO.File]::WriteAllText($file, $content, [System.Text.UTF8Encoding]::new($true))
```

---

## Einschränkungen

- **Kommentare** werden in Exchange-Gruppen-Postfächern gespeichert und sind über die
  Planner API nicht zugänglich
- **Dateianhänge** werden nur als Link-Referenzen exportiert, nicht die Dateien selbst
  (diese liegen in SharePoint)
- **Aufgabenverläufe** (wer hat wann was geändert) werden nicht exportiert
- **Rate Limits**: Microsoft Graph hat Begrenzungen von ~2000 Requests/Minute.
  Das Script beinhaltet automatisches Throttling und Retry-Logik.

---

## Tests

This project includes a comprehensive test suite with 59 unit tests covering both Export and Import functionality.

### Running Tests

```powershell
# Quick test run
Invoke-Pester -Path ./tests

# Or use the test runner with detailed output
pwsh ./tests/Run-Tests.ps1 -Detailed
```

### Test Coverage

- **Export-PlannerData.ps1**: 21 tests covering logging, data export, file handling, and error scenarios
- **Import-PlannerData.ps1**: 38 tests covering import logic, user mapping, dry-run mode, and data restoration
- **Integration Tests**: Manual test scenarios documented for real-world validation

For detailed information, see:

- [tests/README.md](tests/README.md) - Test documentation and setup
- [tests/USAGE.md](tests/USAGE.md) - Practical examples and CI/CD integration
- [tests/Integration-Tests.ps1](tests/Integration-Tests.ps1) - Manual testing scenarios

---

## Lizenz & Support

Created by Alexander Waller, February 2026.

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
