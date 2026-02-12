# Microsoft Planner Export/Import Tool

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-orange)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

---

> ⚠️ **WICHTIG:** Bitte lesen Sie die [WARNING.md](WARNING.md) für wichtige Hinweise zu Haftungsausschluss und Testumgebung!

---

## Übersicht

Dieses Toolset ermöglicht den fast vollständigen Export und Import von Microsoft Planner-Daten
über die Microsoft Graph API. Es wurde speziell für den Lizenzwechsel erstellt, um
Datenverluste zu vermeiden.

### Was wird exportiert/importiert?

| Datenpunkt | Export | Import |
|---|---|---|
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

### 1. PowerShell 7+ (empfohlen)
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

#### Alle eigenen Pläne exportieren:
```powershell
.\Export-PlannerData.ps1
```

> Standardmäßig wird nach `C:\planner-data\PlannerExport_YYYYMMDD_HHMMSS` exportiert.

#### Export in bestimmtes Verzeichnis:
```powershell
.\Export-PlannerData.ps1 -ExportPath "C:\Backup\Planner"
```

#### Nur bestimmte Gruppen exportieren:
```powershell
.\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."
```

> **Tipp:** Die Gruppen-ID findet man in der URL wenn man die Gruppe in Outlook/Teams öffnet,
> oder über das Azure AD Portal unter Gruppen.

#### Abgeschlossene Tasks einbeziehen:
```powershell
.\Export-PlannerData.ps1 -IncludeCompletedTasks
```

### Import

#### Alle exportierten Pläne importieren (gleiche Gruppen):
```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000"
```

#### In eine bestimmte Gruppe importieren:
```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -TargetGroupId "neue-gruppe-id"
```

#### Probelauf (Dry Run) - zeigt was gemacht würde:
```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -DryRun
```

#### Ohne Zuweisungen importieren:
```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipAssignments
```

#### Abgeschlossene Tasks beim Import überspringen:
```powershell
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipCompletedTasks
```

#### Benutzer-Mapping (wenn User-IDs sich ändern):
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

```
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
   # Alle Pläne exportieren
   .\Export-PlannerData.ps1 -ExportPath "C:\Backup\Planner_PreMigration"
   
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
|---|---|
| "Insufficient privileges" | Azure AD Admin muss die App-Berechtigungen freigeben |
| "429 Too Many Requests" | Script wartet automatisch, ggf. `-ThrottleDelayMs 1000` erhöhen |
| Zuweisungen fehlen | Benutzer existieren nicht im Tenant → `-SkipAssignments` oder UserMapping |
| Leerer Export | Prüfe ob der Account Planner-Lizenz hat und Mitglied der Gruppen ist |
| Kommentare fehlen | Kommentare sind über die API nicht exportierbar (Exchange-basiert) |

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

## Lizenz & Support

Created by Alexander Waller, February 2026.

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
