# ⚠️ WARNUNG / WARNING

## Haftungsausschluss

**Wichtige Hinweise zur Verwendung dieses Tools:**

1. **Keine Gewährleistung**
   Das Tool wird ohne Gewähr übergeben. Es gibt keine Garantie für Vollständigkeit, Korrektheit oder Fehlerfreiheit.

2. **Verantwortung des Dienstleisters**
   Der Dienstleister ist vollständig verantwortlich für die korrekte Konfiguration, einschließlich:
   - App Registration in Azure AD
   - Berechtigungen (Permissions) für Microsoft Graph API
   - Tenant-ID und Gruppen-Konfiguration
   - Benutzer-Mapping bei Tenant-Migration

3. **Keine Haftung**
   Der Autor übernimmt keine Haftung für:
   - Datenverlust
   - Fehlerhafte Migration
   - Schäden durch unsachgemäße Verwendung
   - Produktionsausfälle

4. **Testumgebung erforderlich**
   ⚠️ **WICHTIG:** Der Dienstleister soll das Tool zuerst in einer Testumgebung ausführen, bevor es in der Produktivumgebung eingesetzt wird.

5. **Backup erforderlich**
   Erstellen Sie vor der Verwendung des Tools immer ein vollständiges Backup aller Daten.

6. **API-Limitierungen - Daten die NICHT migriert werden können**
   ⚠️ **Folgende Daten können technisch NICHT exportiert/importiert werden:**
   - **Kommentare** (werden in Exchange-Gruppen-Postfächern gespeichert, nicht über Planner API zugänglich)
   - **Dateianhänge** (nur Link-Referenzen werden migriert, nicht die Dateien selbst - diese liegen in SharePoint)
   - **Aufgabenverläufe** (Änderungshistorie: wer hat wann was geändert)

7. **Keine Rollback-Funktion**
   Es gibt KEINE automatische Rückgängig-Funktion. Einmal importierte Daten können nicht automatisch gelöscht werden. Sie müssen manuell gelöscht werden.

8. **Benutzer-Benachrichtigungen**
   Importierte Tasks können Benachrichtigungen an zugewiesene Benutzer auslösen. Benutzer sehen möglicherweise neue Tasks in ihren Planner-Benachrichtigungen.

9. **Microsoft Graph API Rate Limits**
   - Microsoft limitiert API-Anfragen auf ca. 2000 Requests/Minute
   - Das Tool enthält automatisches Throttling und Retry-Logik
   - Bei großen Datenmengen kann die Migration mehrere Stunden dauern
   - Fehlgeschlagene Requests werden automatisch wiederholt

10. **Timing und Ausführungszeitpunkt**
    - Führen Sie die Migration **NICHT während der Geschäftszeiten** aus
    - Planen Sie ein Wartungsfenster ein (z.B. Wochenende oder nach Feierabend)
    - Benutzer sollten während der Migration nicht aktiv in Planner arbeiten

11. **Datenschutz und DSGVO bei Cross-Tenant Migration**
    ⚠️ Bei Migrationen zwischen verschiedenen Tenants:
    - Prüfen Sie die rechtlichen Voraussetzungen (DSGVO, Datenschutz)
    - Holen Sie ggf. die Zustimmung der Betroffenen ein
    - Dokumentieren Sie die Verarbeitungstätigkeiten
    - Prüfen Sie Auftragsverarbeitungsverträge (AVV)

---

## Disclaimer (English)

**Important notes on using this tool:**

1. **No Warranty**
   The tool is provided without warranty. There is no guarantee of completeness, correctness, or freedom from errors.

2. **Service Provider Responsibility**
   The service provider is fully responsible for proper configuration, including:
   - App Registration in Azure AD
   - Permissions for Microsoft Graph API
   - Tenant-ID and group configuration
   - User mapping for tenant migrations

3. **No Liability**
   The author assumes no liability for:
   - Data loss
   - Faulty migration
   - Damage caused by improper use
   - Production outages

4. **Test Environment Required**
   ⚠️ **IMPORTANT:** The service provider must first test the tool in a test environment before using it in production.

5. **Backup Required**
   Always create a complete backup of all data before using the tool.

6. **API Limitations - Data that CANNOT be migrated**
   ⚠️ **The following data CANNOT be exported/imported technically:**
   - **Comments** (stored in Exchange group mailboxes, not accessible via Planner API)
   - **File attachments** (only link references are migrated, not the files themselves - they remain in SharePoint)
   - **Task history** (change audit trail: who changed what and when)

7. **No Rollback Function**
   There is NO automatic undo function. Once imported, data cannot be automatically deleted. Manual deletion is required.

8. **User Notifications**
   Imported tasks may trigger notifications to assigned users. Users may see new tasks in their Planner notifications.

9. **Microsoft Graph API Rate Limits**
   - Microsoft limits API requests to approximately 2000 requests/minute
   - The tool includes automatic throttling and retry logic
   - For large datasets, migration may take several hours
   - Failed requests are automatically retried

10. **Timing and Execution Schedule**
    - Do **NOT run the migration during business hours**
    - Schedule a maintenance window (e.g., weekend or after hours)
    - Users should not actively work in Planner during migration

11. **Data Privacy and GDPR for Cross-Tenant Migration**
    ⚠️ For migrations between different tenants:
    - Review legal requirements (GDPR, data protection)
    - Obtain consent from affected parties if necessary
    - Document processing activities
    - Review data processing agreements (DPA)

---

**Bei Fragen oder Problemen erstellen Sie bitte ein Issue auf GitHub.**
**For questions or problems, please create an issue on GitHub.**
