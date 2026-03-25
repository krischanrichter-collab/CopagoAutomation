# CopagoAutomation - Projektstatus

## Aktueller Stand (24. März 2026)

*   **Refactoring abgeschlossen:** Die Architektur wurde erfolgreich auf relative Koordinaten umgestellt. Die `WindowAutomation`-Klasse nutzt nun `ClientToScreen` und `GetClientRect`, um Klicks relativ zum inneren Bereich des Copago-Fensters auszuführen.
*   **Kompilierungsfehler behoben:** Alle Syntax- und Referenzfehler, die durch das Refactoring entstanden sind, wurden behoben. Das Projekt lässt sich fehlerfrei kompilieren.
*   **UI-Anpassungen:** Die `MainWindow.xaml` und `MainWindow.xaml.cs` wurden aktualisiert, um die neuen Speicher-Modi (Semco Upload vs. Alternativer Ordner) zu unterstützen. Fehlende Event-Handler wurden hinzugefügt.
*   **Kalibrierung:** Die Kalibrierungslogik wurde angepasst, um relative Koordinaten zu speichern, wenn ein Fenster gebunden ist.

## Nächste Schritte

1.  **Test der relativen Klicks:** Der Benutzer muss die Anwendung lokal testen, eine neue Kalibrierung durchführen und prüfen, ob die Klicks nun an den richtigen Stellen landen, unabhängig von der Fensterposition.
2.  **Implementierung der SaveDialog-Automatisierung:**
    *   Die `WindowAutomation`-Klasse muss erweitert werden, um den "Speichern unter"-Dialog von Windows zu erkennen und zu steuern.
    *   Dies beinhaltet das Setzen des Dateipfads im entsprechenden Textfeld und das Klicken auf den "Speichern"-Button.
3.  **Implementierung des PathResolvers:**
    *   Die Logik zur Generierung der korrekten Dateipfade basierend auf dem gewählten Modus (Semco vs. Alternativ), dem Report-Typ (ABC vs. X-Liste) und dem POS muss implementiert werden.
    *   Integration des `PathResolver` in die `AbcAutomation` und `XAutomation`.
4.  **Integration von SaveDialog und PathResolver:**
    *   Verbindung der generierten Pfade mit der SaveDialog-Automatisierung in den Haupt-Automatisierungsschleifen.

## Architektur-Entscheidungen

*   **Relative Koordinaten:** Um die Automatisierung robuster gegen Fensterverschiebungen zu machen, werden Koordinaten relativ zur oberen linken Ecke des Client-Bereichs (`ClientRect`) des Copago-Fensters gespeichert und berechnet. Es ist wichtig zu beachten, dass die Kalibrierungspunkte **nicht** in Abhängigkeit zum Kalibrierungs-Dialogfenster gemappt werden, sondern immer zum Hauptfenster der Copago-Anwendung.
*   **Trennung von UI und Logik:** Die Automatisierungslogik (`AbcAutomation`, `XAutomation`) ist von der UI getrennt und wird über den `AutomationService` aufgerufen.
*   **Zentrale Konfiguration:** Einstellungen werden in `AppSettings` gespeichert und über den `SettingsStore` verwaltet.

## Bekannte Probleme / Offene Punkte

*   Die Event-Handler für Start/Stop/Reset in `MainWindow.xaml.cs` sind teilweise noch leer und müssen mit der entsprechenden Logik aus dem `MainViewModel` oder `AutomationService` verknüpft werden.
*   Die Automatisierung für den Save-Dialog ist noch auskommentiert (`// TODO: SaveDialog Automation hier einfügen`).

## Sicherheitshinweis

*   **GitHub Token:** Der Personal Access Token (PAT) für GitHub ist im Git-Credential-Manager der Sandbox gespeichert. Er darf **niemals** im Klartext in den Projektdateien abgelegt werden.
