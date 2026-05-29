# Funktionsreferenz

Übersicht aller Funktionen in `Export-PlannerData.ps1` und `Import-PlannerData.ps1`.

---

## Export-PlannerData.ps1

### Parameter

| Parameter | Typ | Beschreibung |
| --- | --- | --- |
| `-ExportPath` | String | Zielverzeichnis. Standard: `C:\temp`, `C:\tmp` oder Dokumente |
| `-GroupIds` | String[] | M365-Gruppen-IDs, aus denen exportiert wird |
| `-GroupNames` | String[] | M365-Gruppennamen (Teilstring-Suche) |
| `-Interactive` | Switch | Interaktive Gruppenauswahl aus Liste aller verfügbaren Gruppen |
| `-UseCurrentUser` | Switch | Exportiert alle Pläne des angemeldeten Benutzers |
| `-IncludeCompletedTasks` | Switch | Exportiert auch abgeschlossene Tasks (Standard: nur aktive) |
| `-TenantId` | String | Azure AD Tenant-ID. Wird nach erstem Login gespeichert |

### Funktionen

#### `Write-PlannerLog`

Gibt eine formatierte Log-Zeile mit Zeitstempel und Level aus. Wird automatisch per `Start-Transcript` in die Log-Datei geschrieben.

```
[$timestamp] [$Level] $Message
```

Level: `INFO` (weiß), `OK` (grün), `WARN` (gelb), `ERROR` (rot)

---

#### `Test-SafePath`

Validiert einen Dateisystempfad auf Sicherheit und Zugänglichkeit vor dem Export oder Import.

Prüfungen:
- Leerer Pfad
- UNC-Pfade (Netzwerkpfade `\\server\...`) — blockiert aus Sicherheitsgründen
- Pfad normalisieren (`..` auflösen)
- Export-Modus: Schreibrechte auf Zielverzeichnis
- Import-Modus: Existenz und Leserechte auf Quellverzeichnis

---

#### `Connect-ToGraph`

Authentifiziert gegen Microsoft Graph API mit den Scopes:
`Group.Read.All`, `Tasks.Read`, `Tasks.ReadWrite`, `User.Read`, `User.ReadBasic.All`

Ablauf:
1. Prüft ob bereits eine aktive Session besteht (`Get-MgContext`)
2. Liest gespeicherte Anmeldedaten aus `~/.planner-auth.json`
3. Versucht Silent-Auth über MSAL-Cache (kein Browser)
4. Falls nötig: Interaktive Anmeldung — Benutzer wählt Browser oder Device-Code
5. Speichert Tenant-ID und Konto in `~/.planner-auth.json` für nächste Ausführung

---

#### `Get-AllM365Groups`

Ruft alle M365-Gruppen (Unified Groups) aus dem Tenant ab. Unterstützt Paging für große Tenants. Ergebnis wird alphabetisch sortiert zurückgegeben.

---

#### `Get-GroupsByNames`

Sucht M365-Gruppen anhand von Anzeigenamen. Unterstützt exakte Übereinstimmung und Präfix-Suche (`startswith`). Gibt nur Unified Groups zurück.

---

#### `Show-GroupSelectionMenu`

Zeigt eine nummerierte Liste verfügbarer Gruppen im Terminal an. Benutzer kann einzelne Nummern (`1,3,5`), `A` für alle Gruppen, oder `0` zum Abbrechen eingeben.

---

#### `Get-AllUserPlans`

Lädt alle Planner-Pläne des angemeldeten Benutzers über zwei Methoden:
1. Über die Gruppen des Benutzers (`/me/memberOf/.../planner/plans`)
2. Direkt über `/me/planner/plans` (als Ergänzung)

Duplikate werden herausgefiltert.

---

#### `Get-PlansByGroupIds`

Lädt alle Planner-Pläne für eine Liste von Gruppen-IDs. Verarbeitet 403-Fehler (fehlende Mitgliedschaft) separat mit hilfreicher Fehlermeldung. Erkennt Doppelplanenamen innerhalb des Laufs und warnt im Log.

Gibt immer ein Array zurück (verhindert PowerShell-Pipeline-Unrolling bei leerem Ergebnis mit `,`-Operator).

---

#### `Export-PlanDetails`

Hauptfunktion des Exports. Lädt für einen Plan:
- Plan-Details (Kategoriebeschreibungen/Labels, ETag)
- Buckets
- Tasks (mit Paging via `@odata.nextLink`)
- Task-Details (Beschreibungen, Checklisten, Referenzen) — mit 200ms Pause zwischen Requests (Rate Limiting)
- Benutzerinformationen für alle Zuweisungen (DisplayName, UPN, Mail)

Schreibt das Ergebnis als JSON-Datei. Warnt wenn zwei Pläne im selben Lauf denselben Dateinamen erzeugen.

---

#### `Export-ReadableSummary`

Erstellt eine lesbare Textdatei (`*_Zusammenfassung.txt`) aus den exportierten Plan-Daten mit:
- Plan- und Gruppenname, Exportdatum
- Kategorien/Labels des Plans
- Alle Buckets mit ihren Tasks (Status, Priorität, Fälligkeitsdatum, Startdatum, Zuweisungen, Labels, Beschreibung, Checkliste, Links)
- Tasks ohne Bucket
- Statistik (Gesamt, nicht begonnen, in Bearbeitung, abgeschlossen)

---

### Hauptprogramm-Ablauf (Export)

```
1. Transcript starten (logs/export_YYYYMMDD_HHMMSS.log)
2. Export-Verzeichnis ermitteln / validieren / erstellen
3. Microsoft.Graph Modul prüfen / laden
4. Connect-ToGraph
5. Pläne laden (je nach Parameter):
   -UseCurrentUser  → Get-AllUserPlans
   -GroupIds        → Get-PlansByGroupIds
   -GroupNames      → Get-GroupsByNames + Get-PlansByGroupIds
   -Interactive     → Get-AllM365Groups + Show-GroupSelectionMenu + Get-PlansByGroupIds
   (kein Parameter) → Get-AllM365Groups + Frage: Auswahl oder alle
6. Für jeden Plan: Export-PlanDetails + Export-ReadableSummary
7. _ExportIndex.json schreiben
8. Abschlussmeldung
9. Disconnect-MgGraph
```

---

## Import-PlannerData.ps1

### Parameter

| Parameter | Typ | Beschreibung |
| --- | --- | --- |
| `-ImportPath` | String | Verzeichnis mit exportierten JSON-Dateien (Pflicht) |
| `-TargetGroupId` | String | Zielgruppe für alle Pläne (überschreibt Original-Gruppe) |
| `-DryRun` | Switch | Probelauf ohne Änderungen |
| `-SkipAssignments` | Switch | Zuweisungen nicht importieren |
| `-SkipCompletedTasks` | Switch | Abgeschlossene Tasks (100%) nicht importieren |
| `-UserMapping` | Hashtable | Manuelle User-ID-Zuordnung: `@{"alt-id"="neu-id"}` |
| `-TenantId` | String | Azure AD Tenant-ID |

### Funktionen

#### `Write-PlannerLog` / `Write-Log`

Wie im Export, zusätzlich mit Level `DRYRUN` (magenta) für Probelauf-Ausgaben.

---

#### `Invoke-GraphWithRetry`

Wrapper um `Invoke-MgGraphRequest` mit:
- Automatischer Retry-Logik bei 429 (Too Many Requests) — liest `Retry-After`-Header
- Exponentielles Backoff bei anderen Fehlern
- Unterstützt `Method`, `Uri`, `Body`, `Headers` Parameter
- Standardmäßig 3 Versuche

---

#### `Add-ErrorToTracker`

Fügt einen Fehler zur zentralen Fehlerstatistik (`$script:errorTracker`) hinzu. Wird am Ende der Ausführung als Zusammenfassung ausgegeben.

Kategorien: `Plans`, `Buckets`, `Tasks`, `TaskDetails`, `UserResolution`

---

#### `ConvertTo-IsoDate`

Konvertiert Datumswerte aus dem Export in ISO 8601 Format (`yyyy-MM-ddTHH:mm:ssZ`).

Unterstützt:
- `System.DateTime` — wird direkt nach UTC konvertiert (PowerShell 7 wandelt `/Date(ms)/` automatisch um)
- `System.DateTimeOffset`
- OData v3 Format `/Date(1234567890000)/` (Unix-Millisekunden)
- ISO-Strings (locale-unabhängig über `InvariantCulture`)

---

#### `Get-OrCreateWarningCategory`

Findet den höchsten freien Kategorie-Slot im Plan (category25 bis category1) und legt dort eine **"⚠ Import-Fehler"**-Kategorie an. Gibt den Slot-Schlüssel zurück (z.B. `"category25"`).

Wird lazy initialisiert — nur wenn tatsächlich ein Problem bei einem Task auftritt.

---

#### `Get-ExistingPlansForGroup`

Lädt alle bestehenden Planner-Pläne einer Gruppe (mit Caching pro Gruppen-ID), um Namenskonflikte beim Import zu erkennen.

---

#### `Resolve-UserId`

Löst eine alte User-ID aus dem Export in eine ID im Ziel-Tenant auf. Strategie:
1. UserMapping-Hashtable prüfen
2. Lookup per UserPrincipalName im Ziel-Tenant
3. Lookup per Mail-Adresse (falls abweichend von UPN)
4. Originale ID direkt versuchen (Same-Tenant-Szenario)

---

#### `Import-PlanFromJson`

Hauptfunktion des Imports. Verarbeitet eine exportierte JSON-Datei:

1. Zielgruppe ermitteln (Parameter, Original-Gruppe oder Abfrage)
2. **DryRun:** Zeigt Vorschau ohne Änderungen
3. Plan erstellen (oder bestehenden verwenden/überschreiben/überspringen bei Namenskonflikt)
4. Kategorien/Labels setzen (PATCH mit ETag)
5. Buckets erstellen (ID-Mapping alte → neue ID)
6. Tasks erstellen mit:
   - Bucket-Zuordnung
   - Fälligkeits-/Startdatum (Validierung: Start ≤ Fällig)
   - Labels/Kategorien
   - Zuweisungen (via `Resolve-UserId`)
7. Task-Details setzen (Beschreibung, Checkliste, Referenzen) — PATCH mit ETag
8. Bei Problemen: Task mit "⚠ Import-Fehler"-Label markieren
9. Import-Mapping als JSON speichern (`*_ImportMapping.json`)

---

#### `Get-M365GroupsForListing`

Lädt alle M365-Gruppen aus dem Tenant für die interaktive Zielgruppenauswahl beim Import.

---

#### `Find-PlannerExports`

Durchsucht Standard-Speicherorte nach Planner-Export-Verzeichnissen (`PlannerExport_*`).

---

### Hauptprogramm-Ablauf (Import)

```
1. Transcript starten (logs/import_YYYYMMDD_HHMMSS.log)
2. Import-Verzeichnis validieren
3. Microsoft.Graph Modul prüfen / laden
4. Connect-ToGraph (mit erweiterten Scopes: Group.ReadWrite.All, Tasks.ReadWrite)
5. JSON-Dateien im Import-Verzeichnis laden
6. Für jede JSON-Datei: Import-PlanFromJson
7. Fehlerstatistik ausgeben
```

---

## Ausgabedateien

| Datei | Erstellt von | Inhalt |
| --- | --- | --- |
| `<Plan>.json` | Export | Vollständige Plandaten für Import |
| `<Plan>_Zusammenfassung.txt` | Export | Lesbare Textübersicht |
| `_ExportIndex.json` | Export | Metadaten: Datum, Konto, alle Pläne mit Task-Anzahl |
| `<Plan>_ImportMapping.json` | Import | ID-Zuordnung alte→neue Plan/Bucket/Task-IDs |
| `logs/export_*.log` | Export | Vollständiges Protokoll des Export-Laufs |
| `logs/import_*.log` | Import | Vollständiges Protokoll des Import-Laufs |

---

## Bekannte Einschränkungen

| Einschränkung | Grund |
| --- | --- |
| Kommentare nicht exportierbar | In Exchange gespeichert, kein Zugriff über Planner API |
| Dateianhänge nur als Link | Dateien liegen in SharePoint, werden nicht übertragen |
| Aufgabenverlauf nicht exportierbar | Kein API-Endpunkt verfügbar |
| Kein Zugriff ohne Gruppenmitgliedschaft | Planner API erzwingt Mitgliedschaft (403 Forbidden) |
| Tenant-spezifische orderHints | Werden beim Import nicht übernommen, da im Ziel-Tenant ungültig |
| Maximal 25 Labels pro Plan | Planner-Limitierung (category1–category25) |
