# Outlook Gateway Integration Checklist

This guide outlines the tasks required to connect Microsoft Outlook clients to an internal Red Ink gateway in enterprise environments. Each section expands on the high-level requirements so administrators can prepare networking, authentication, policy, and offline workflows before rolling out the add-in.

## 1. Netzwerk- und Firewall-Regeln
- Öffnen Sie die notwendigen Ports zwischen den Outlook-Clients und dem internen Gateway (z. B. HTTPS auf TCP 443).
- Whitelisten Sie die FQDNs bzw. IP-Subnetze des Gateways in allen Zwischenfirewalls, Web-Proxies und Endpoint-Firewalls.
- Aktivieren Sie, falls erforderlich, NTLM- oder Kerberos-Passthrough, damit Outlook vorhandene Windows-Anmeldetickets zum Gateway weiterreichen kann.
- Dokumentieren Sie etwaige SSL/TLS-Inspection oder DPI-Geräte; konfigurieren Sie Ausnahmen, damit MAPI/HTTP- oder EWS-Sitzungen nicht unterbrochen werden.
- Verifizieren Sie die Konfiguration mit Test-Clients (z. B. `Test-EmailAutoConfig`) und überwachen Sie Netzwerk-Logs auf blockierte Anfragen.

## 2. Single Sign-on (SSO)
- Entscheiden Sie, ob EWS oder MAPI/HTTP die maßgebliche Authentifizierungsquelle für den Benutzerkontext ist.
- Implementieren Sie auf dem Gateway eine Vertrauenskette, die Outlook-Tokens übernimmt (Kerberos-Delegation oder NTLM-Constrained Delegation, falls erforderlich).
- Leiten Sie erhaltene EWS/MAPI-Tokens serverseitig weiter und validieren Sie deren Gültigkeit (Zeitstempel, SPN, Signatur).
- Ergänzen Sie Fallback-Mechanismen (z. B. OAuth 2.0 Device Code) für Fälle, in denen kein Domänenkontext vorliegt.
- Testen Sie SSO mit einem Domänen- und einem Nicht-Domänenkonto und protokollieren Sie Fehlversuche.

## 3. Lokale Gruppenrichtlinien (GPO)
- Pflegen Sie eine Whitelist für erforderliche Outlook-Add-ins (Red Ink) über Administrative Templates (Office > Add-ins).
- Legen Sie eine Automatik für Updates fest (internes Fileshare, SCCM oder Microsoft Endpoint Manager) und dokumentieren Sie den Update-Kanal.
- Konfigurieren Sie Signaturüberprüfungen und eventuelle "LoadBehavior"-Werte, um das Add-in bei Start zu aktivieren.
- Richten Sie Richtlinien für vertrauenswürdige Speicherorte und Makroeinstellungen ein, falls das Add-in Skripte oder VBA benötigt.
- Halten Sie ein Rollback-Skript bereit, das die Richtlinien bei Problemen zurücknimmt.

## 4. Offline-Modus und Fallback
- Definieren Sie, wie das Add-in Anfragen puffert, wenn das Gateway nicht erreichbar ist (lokale Warteschlangen oder temporäre Speicherung).
- Stellen Sie dem Nutzer eine klare Statusanzeige in Outlook bereit (z. B. InfoBar oder Taskpane-Hinweis), falls das Gateway offline ist.
- Dokumentieren Sie, wie Benutzer Offline-Antworten überprüfen und nachträglich senden können, sobald die Verbindung wieder besteht.
- Planen Sie einen Support-Prozess, der Nutzer proaktiv über geplante Downtimes oder Störungen informiert (E-Mail, Teams, ServiceNow).
- Führen Sie regelmäßige Wiederherstellungstests durch, um sicherzustellen, dass gepufferte Inhalte nach Wiederherstellung der Verbindung verarbeitet werden.

## Prüf- und Abnahmeschritte
1. Testfall für jede Firewall-Regel (Inbound/Outbound) durchführen und protokollieren.
2. SSO mit unterschiedlichen Authentifizierungsarten verifizieren (Domänenkonto, Externes Konto, Fallback).
3. GPO-Verteilung in einer Pilot-OU anwenden und auf Compliance kontrollieren.
4. Offline-Szenario in einer Testumgebung simulieren und beobachten, wie das Add-in Nutzer informiert und Warteschlangen abarbeitet.
5. Ergebnisse dokumentieren und den Betrieb an den IT-Service-Desk übergeben.
