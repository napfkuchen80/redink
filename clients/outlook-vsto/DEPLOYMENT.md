# Bereitstellung des Outlook Add-ins

Dieses Dokument beschreibt das empfohlene Deployment über ClickOnce sowie Alternativen per MSI.

## ClickOnce-Bereitstellung

1. **Voraussetzungen prüfen**
   - Stellen Sie sicher, dass ein Code Signing Zertifikat (mindestens Code Signing SHA-2) vorliegt.
   - Aktivieren Sie in Visual Studio unter *Projekt ▸ Eigenschaften ▸ Signierung* die Option **"Manifeste signieren"** und wählen Sie das Zertifikat aus.
2. **Publikationsprofil anlegen**
   - Öffnen Sie im Projektmappen-Explorer das Projekt `RedInk.OutlookAddIn`.
   - Wählen Sie *Veröffentlichen…* und konfigurieren Sie einen Zielordner (z. B. eine Netzfreigabe oder HTTPS-Webserver).
   - Aktivieren Sie automatische Updates (z. B. Intervall 7 Tage) und geben Sie die Update-URL an.
3. **Build & Veröffentlichung**
   - Erstellen Sie einen Release-Build.
   - Veröffentlichen Sie über das angelegte Profil. Visual Studio erzeugt dabei die Setup.exe, die Application Files sowie die Manifeste.
4. **Installation auf Clients**
   - Signieren Sie optional die Setup.exe zusätzlich mit `signtool.exe`.
   - Verteilen Sie die Setup.exe sowie den `application`-Ordner auf den Zielsystemen.
   - Anwender führen die Setup.exe mit erhöhten Rechten aus. Nach der Installation verwaltet ClickOnce künftige Updates selbständig.

## Bereitstellung per MSI (Enterprise Szenario)

1. **WiX Toolset installieren** oder ein alternatives Authoring-Tool (z. B. Advanced Installer).
2. **VSTO Runtime** als Voraussetzung in das MSI aufnehmen (`vstor_redist.exe`).
3. **Primäre Ausgabe** des Projekts (`Primary Output`) sowie den `app.config` in das MSI integrieren.
4. **Custom Actions** hinzufügen, die das Add-in im Office-Vertrauensspeicher registrieren (`AddIn`-Eintrag in `HKCU\Software\Microsoft\Office\Outlook\Addins`).
5. **Digital signieren** des MSI-Pakets mit dem Unternehmenszertifikat.
6. **Verteilung** über Gruppenrichtlinien, Intune oder SCCM.

## Vertrauenswürdigkeit & Sicherheit

- Aktivieren Sie die Signierung des VSTO-Manifests, damit Outlook die Quelle verifizieren kann.
- Hinterlegen Sie das Zertifikat in den vertrauenswürdigen Herausgebern der Windows-Clients (z. B. per GPO).
- Kommunizieren Sie die Add-in-Quelle in Outlook unter *Datei ▸ Optionen ▸ Trust Center ▸ Vertrauenswürdige Add-ins*.
- Hinterlegen Sie die Basis-URL des ClickOnce-Deployments in den vertrauenswürdigen Speicherorten, wenn eine Webfreigabe genutzt wird.

## Konfigurationsdaten

- API- und Gateway-Einstellungen werden aus `app.config` oder den Umgebungsvariablen `LLM_GATEWAY_BASEURL` und `LLM_GATEWAY_APIKEY` geladen.
- Für produktive Deployments sollte das API-Secret per Windows Credential Manager oder Intune Secrets bereitgestellt werden. Entfernen Sie den Platzhalterwert aus `app.config`, bevor Sie das Paket signieren.

## Update- und Rollback-Strategie

- ClickOnce speichert automatisch die vorangegangene Version. Über das Windows-Startmenü kann der Benutzer bei Bedarf auf die letzte funktionierende Version zurückrollen.
- Bei MSI-Rollouts sollte eine separate `x.y.z`-Upgrade-Strategie festgelegt und bei Fehlern ein Downgrade-Paket bereitgestellt werden.

## Logging

- Aktivieren Sie in Outlook die VSTO-Logdateien (`HKCU\Software\Microsoft\Office\Outlook\Addins\LoggingLevel = 1`) für Pilotgruppen.
- Ergänzen Sie nach Bedarf Application Insights oder ein internes Telemetrie-Endpunkt im `LlmGatewayClient`.
