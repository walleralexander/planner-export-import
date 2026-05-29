# Microsoft Planner Export/Import Tool

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-orange)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

Erstellt von Alexander Waller · Hohenems Informatik

---

## Übersicht

Dieses Tool ermöglicht den vollständigen Export und Import von Microsoft Planner-Daten über die Microsoft Graph API. Hauptanwendungsfall ist die Datensicherung vor Lizenzwechseln sowie die Migration von Plänen zwischen M365-Gruppen oder Tenants.

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

\* Zuweisungen funktionieren nur wenn die Benutzer im Ziel-Tenant mit gleicher UPN/Mail existieren
\** Kommentare sind über die Planner API nicht zugänglich (werden in Exchange gespeichert)

---

## Voraussetzungen

### 1. PowerShell

PowerShell 5.1 (Windows vorinstalliert) wird unterstützt. PowerShell 7+ empfohlen:

```powershell
winget install Microsoft.PowerShell
```

### 2. Microsoft Graph PowerShell Module

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### 3. M365-Berechtigungen

Beim ersten Ausführen öffnet sich ein Browser-Fenster (oder Device-Code-Flow). Das angemeldete Konto benötigt:

- **Export:** Mitgliedschaft in den zu exportierenden Gruppen
- **Import:** Mitgliedschaft in den Zielgruppen + Recht zum Erstellen von Plänen
- Admin-Consent für die erforderlichen Graph-Berechtigungen (siehe [CLOUD-SETUP.md](CLOUD-SETUP.md))

---

## Verwendung

### Export

```powershell
# Alle eigenen Pläne exportieren (empfohlen für einfache Fälle)
.\Export-PlannerData.ps1 -UseCurrentUser

# Mit abgeschlossenen Tasks
.\Export-PlannerData.ps1 -UseCurrentUser -IncludeCompletedTasks

# Bestimmte Gruppen nach Name
.\Export-PlannerData.ps1 -GroupNames "Projektteam Alpha", "Marketing Team"

# Bestimmte Gruppen nach ID
.\Export-PlannerData.ps1 -GroupIds "abc123-...", "def456-..."

# Interaktive Gruppenauswahl
.\Export-PlannerData.ps1 -Interactive

# Ohne Parameter: zeigt alle Gruppen, fragt ob Auswahl oder alle exportieren
.\Export-PlannerData.ps1

# In bestimmtes Verzeichnis exportieren
.\Export-PlannerData.ps1 -UseCurrentUser -ExportPath "C:\Backup\Planner"
```

> **Export-Verzeichnis:** Wird automatisch ermittelt: `C:\temp` → `C:\tmp` → Dokumente-Ordner.
> Kann mit `-ExportPath` überschrieben werden.

### Import

```powershell
# Alle exportierten Pläne importieren (in Original-Gruppen)
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000"

# Probelauf (zeigt was gemacht würde, ohne Änderungen)
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -DryRun

# In eine bestimmte Zielgruppe importieren
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -TargetGroupId "neue-gruppe-id"

# Ohne Zuweisungen importieren
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipAssignments

# Abgeschlossene Tasks nicht importieren
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -SkipCompletedTasks

# Mit Benutzer-Mapping (bei Tenant-Migration mit geänderten User-IDs)
$mapping = @{
    "alte-user-id-1" = "neue-user-id-1"
    "alte-user-id-2" = "neue-user-id-2"
}
.\Import-PlannerData.ps1 -ImportPath ".\PlannerExport_20260209_143000" -UserMapping $mapping
```

---

## Exportierte Dateien

```text
PlannerExport_20260209_143000/
├── _ExportIndex.json                        # Metadaten: Datum, Konto, Planliste
├── Projektplanung.json                      # Strukturierte Daten für Import
├── Projektplanung_Zusammenfassung.txt       # Lesbare Textübersicht
├── IT_Infrastruktur.json
├── IT_Infrastruktur_Zusammenfassung.txt
└── ...

logs/
├── export_20260209_143000.log               # Vollständiges Protokoll je Export-Lauf
└── import_20260209_150000.log               # Vollständiges Protokoll je Import-Lauf
```

---

## Empfohlene Vorgehensweise beim Lizenzwechsel

**1. Vor dem Wechsel — Export:**

```powershell
.\Export-PlannerData.ps1 -UseCurrentUser -ExportPath "C:\Backup\Planner_VorWechsel"
# Zusammenfassungen prüfen:
Get-ChildItem "C:\Backup\Planner_VorWechsel\*Zusammenfassung*"
```

**2. Export-Verzeichnis sichern** (Netzlaufwerk, ZIP, externer Speicher)

**3. Nach dem Wechsel — Import:**

```powershell
# Erst Probelauf:
.\Import-PlannerData.ps1 -ImportPath "C:\Backup\Planner_VorWechsel" -DryRun

# Dann tatsächlicher Import:
.\Import-PlannerData.ps1 -ImportPath "C:\Backup\Planner_VorWechsel"
```

**4. Stichproben prüfen:** Pläne in Teams öffnen, Buckets, Tasks, Beschreibungen, Checklisten prüfen

> Tasks mit Import-Problemen (fehlende Zuweisung, ungültige Daten) werden automatisch mit dem Label **"⚠ Import-Fehler"** markiert und sind in Teams leicht filterbar.

---

## Fehlerbehebung

| Problem | Lösung |
| --- | --- |
| 403 Forbidden beim Export | Konto ist kein Mitglied der Gruppe — oder Admin-Consent fehlt (siehe [CLOUD-SETUP.md](CLOUD-SETUP.md)) |
| Zuweisungen fehlen nach Import | Benutzer existieren nicht im Tenant → `-SkipAssignments` oder `-UserMapping` verwenden |
| 429 Too Many Requests | Script wartet automatisch — ggf. Throttle erhöhen |
| Fälligkeitsdaten fehlen | Waren im Export enthalten, werden korrekt importiert (nur wenn StartDate ≤ DueDate) |
| Kommentare fehlen | Über die Planner API nicht exportierbar (Exchange-basiert) |
| Dateianhänge fehlen | Nur Links werden exportiert, Dateien selbst liegen in SharePoint |
| Gruppe fehlt im Export | Konto ist nicht Mitglied → mit `-UseCurrentUser` nur eigene Gruppen exportieren |

---

## Einschränkungen

- **Kommentare** — in Exchange gespeichert, über Planner API nicht zugänglich
- **Dateianhänge** — nur als Link-Referenz, Dateien bleiben in SharePoint
- **Aufgabenverläufe** — Änderungshistorie wird nicht exportiert
- **Rate Limits** — ~2000 Requests/Minute; das Script beinhaltet Throttling und automatischen Retry
- **Gruppen ohne Mitgliedschaft** — können nicht exportiert werden (403 Forbidden)

---

## Weitere Dokumentation

- [CLOUD-SETUP.md](CLOUD-SETUP.md) — M365/Azure AD Einrichtung für Administratoren
- [FUNCTIONS.md](FUNCTIONS.md) — Funktionsreferenz beider Scripts
