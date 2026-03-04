# Cloud-Einrichtung — Microsoft 365 / Azure AD

Diese Anleitung richtet sich an IT-Administratoren und beschreibt, was in Microsoft 365 eingerichtet werden muss, damit das Planner Export/Import Tool funktioniert.

---

## Übersicht

Das Tool verwendet **delegierte Berechtigungen** über die Microsoft Graph API. Das bedeutet:

- Das Tool meldet sich als **Benutzer** an (kein Service Principal / App-Only)
- Es kann nur auf Daten zugreifen, auf die der angemeldete Benutzer selbst Zugriff hat
- Ein Administrator muss einmalig den **Admin-Consent** für die benötigten Berechtigungen erteilen

---

## Schritt 1: Admin-Consent erteilen

Beim ersten Aufruf von `Connect-MgGraph` mit den benötigten Scopes erscheint ein Consent-Dialog. Wenn der Benutzer kein Administrator ist, wird der Zugriff verweigert und ein Admin muss den Consent erteilen.

### Option A: Consent über Browser-Aufruf (empfohlen)

Ein globaler Administrator öffnet folgenden Link und meldet sich an:

```
https://login.microsoftonline.com/<TenantId>/adminconsent?client_id=14d82eec-204b-4c2f-b7e8-296a70dab67e&scope=https://graph.microsoft.com/.default
```

> `14d82eec-204b-4c2f-b7e8-296a70dab67e` ist die Client-ID des offiziellen **Microsoft Graph PowerShell**-Clients, den das Modul `Microsoft.Graph` intern verwendet.

Die `<TenantId>` findet man im Azure Portal unter:
**Azure Active Directory → Übersicht → Verzeichnis-ID**

### Option B: Consent per PowerShell (als Admin)

```powershell
Connect-MgGraph -Scopes "Group.Read.All","Group.ReadWrite.All","Tasks.Read","Tasks.ReadWrite","User.Read","User.ReadBasic.All" -TenantId "<TenantId>"
```

Beim ersten Aufruf erscheint der Consent-Dialog. Als globaler Administrator kann man direkt für die gesamte Organisation zustimmen ("Zustimmung im Namen Ihrer Organisation").

---

## Schritt 2: Benötigte Berechtigungen

Das Tool benötigt ausschließlich **delegierte Berechtigungen** (keine Anwendungsberechtigungen):

| Berechtigung | Export | Import | Beschreibung |
| --- | --- | --- | --- |
| `Group.Read.All` | ✅ | — | Gruppen auflisten und Planner-Pläne einer Gruppe lesen |
| `Group.ReadWrite.All` | — | ✅ | Pläne in Gruppen erstellen |
| `Tasks.Read` | ✅ | — | Planner-Daten lesen (Tasks, Buckets, Details) |
| `Tasks.ReadWrite` | ✅ | ✅ | Planner-Daten lesen und schreiben |
| `User.Read` | ✅ | ✅ | Eigenes Profil lesen (für Authentifizierung) |
| `User.ReadBasic.All` | ✅ | ✅ | Benutzerinformationen für Zuweisungen auflösen |

> Das Export-Script fordert alle Scopes an, damit nach dem Export direkt ein Import möglich ist ohne erneute Anmeldung.

---

## Schritt 3: Benutzeranforderungen

### Gruppenmitgliedschaft

Das angemeldete Konto muss **Mitglied** der M365-Gruppen sein, die exportiert oder importiert werden sollen.

- Planner prüft die Gruppenmitgliedschaft — kein Mitglied = 403 Forbidden
- Auch ein globaler Administrator bekommt 403, wenn er nicht Mitglied der Gruppe ist
- Lösung: Admin temporär zur Gruppe hinzufügen, oder einen Benutzer verwenden, der bereits Mitglied ist

Gruppen verwalten im **Microsoft 365 Admin Center:**
**Gruppen → Aktive Gruppen → Gruppe auswählen → Mitglieder**

### Planner-Lizenz

Der angemeldete Benutzer muss eine Lizenz haben, die Microsoft Planner enthält:

- Microsoft 365 Business Basic / Standard / Premium
- Microsoft 365 E3 / E5
- Office 365 E1 / E3 / E5
- Microsoft Teams Essentials (enthält Planner)

---

## Schritt 4: Authentifizierungs-Cache

Das Tool speichert nach erfolgreicher Anmeldung die **Tenant-ID und den Kontonamen** in:

```
%USERPROFILE%\.planner-auth.json
```

Beim nächsten Start wird versucht, mit diesen Informationen lautlos (ohne Browser) anzumelden. Schlägt das fehl (abgelaufenes Token, geändertes Passwort), erscheint wieder der Login-Dialog.

**Datei manuell löschen** um die gespeicherte Anmeldung zurückzusetzen:

```powershell
Remove-Item "$env:USERPROFILE\.planner-auth.json"
```

---

## Schritt 5: Tenant-Migration (Cross-Tenant Import)

Beim Import in einen anderen Tenant sind zusätzliche Schritte nötig:

### 1. Export im Quell-Tenant

```powershell
.\Export-PlannerData.ps1 -UseCurrentUser -TenantId "<Quell-TenantId>"
```

### 2. Benutzer-IDs mappen

Die User-IDs im Export gehören zum Quell-Tenant und sind im Ziel-Tenant ungültig. Das Import-Script versucht automatisch, Benutzer über UPN/Mail aufzulösen. Wenn UPNs sich geändert haben, muss ein manuelles Mapping angegeben werden:

```powershell
$mapping = @{
    "alte-user-id@quell-tenant.com" = "neue-user-id@ziel-tenant.com"
}
.\Import-PlannerData.ps1 -ImportPath ".\Export" -UserMapping $mapping -TenantId "<Ziel-TenantId>"
```

### 3. Zielgruppen anlegen

Im Ziel-Tenant müssen die M365-Gruppen bereits existieren, bevor importiert wird. Der Import erstellt Pläne innerhalb bestehender Gruppen — er erstellt keine neuen Gruppen.

**Gruppen anlegen im Microsoft 365 Admin Center:**
**Gruppen → Aktive Gruppen → Gruppe hinzufügen → Microsoft 365**

---

## Häufige Fehler und Lösungen

| Fehler | Ursache | Lösung |
| --- | --- | --- |
| 403 Forbidden auf `/groups/{id}/planner/plans` | Konto ist nicht Mitglied der Gruppe | Konto zur Gruppe hinzufügen |
| 403 Forbidden auf `/me/planner/plans` | Planner-Lizenz fehlt | Lizenz prüfen und zuweisen |
| `AADSTS65001: consent required` | Admin-Consent fehlt | Schritt 1 dieser Anleitung durchführen |
| `AADSTS50076: MFA required` | Multi-Factor-Authentication erzwungen | MFA im Browser oder Authenticator abschließen |
| `authentication_canceled` | Benutzer hat Login abgebrochen | Erneut starten, Login durchführen |
| Zuweisungen fehlen nach Cross-Tenant-Import | UPNs haben sich geändert | `-UserMapping` Parameter verwenden |

---

## Checkliste für den Administrator

Vor dem ersten Einsatz:

- [ ] Admin-Consent für Microsoft Graph PowerShell erteilt (Schritt 1)
- [ ] Benutzer hat Microsoft 365 Lizenz mit Planner
- [ ] Benutzer ist Mitglied aller zu exportierenden Gruppen
- [ ] Bei Import: Zielgruppen im Ziel-Tenant angelegt
- [ ] Bei Import: Benutzer ist Mitglied der Zielgruppen
- [ ] Bei Cross-Tenant-Migration: User-Mapping vorbereitet (wenn UPNs abweichen)
