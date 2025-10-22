# Red Ink Outlook VSTO Add-in

Dieses Verzeichnis enthält das neue Outlook VSTO Add-in, das KI-Funktionen für Antwortgenerierung, Zusammenfassungen und Entwurfsprüfungen in Microsoft Outlook bereitstellt.

## Entwicklungsumgebung einrichten

1. **Visual Studio installieren**
   - Verwenden Sie Visual Studio 2022 Professional oder Enterprise.
   - Wählen Sie bei der Installation die Workloads **"Office-/SharePoint-Entwicklung"** sowie **".NET-Desktopentwicklung"** aus.
   - Aktivieren Sie zusätzlich das individuelle Komponentenpaket **"Microsoft Office Developer Tools"**.
2. **Office-Integration vorbereiten**
   - Stellen Sie sicher, dass Microsoft 365 Apps (Outlook) in der gleichen Bitness wie Visual Studio installiert sind.
   - Öffnen Sie Outlook mindestens einmal, damit notwendige Registrierungen vorgenommen werden.
3. **Projekt öffnen**
   - Öffnen Sie die Lösung [`OutlookAddIn.sln`](OutlookAddIn.sln) in Visual Studio.
   - Akzeptieren Sie ggf. die Wiederherstellung der NuGet-Pakete für Office Tools.
4. **Debuggen**
   - Setzen Sie `RedInk.OutlookAddIn` als Startprojekt.
   - Starten Sie das Debugging mit `F5`. Visual Studio öffnet Outlook im Debugmodus und lädt das Add-in automatisch.

Weitere Details zur Bereitstellung finden Sie in [DEPLOYMENT.md](DEPLOYMENT.md).
