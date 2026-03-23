CopagoAutomation – PROJECT_STATUS.md

Last Update: 2026-03-17
Session: Automation Engine Phase 3 – Save Strategy Refactor (Semco Upload + Alternativ Mode)

Projektziel

Die Anwendung automatisiert Report-Workflows in CopagoOffice.

Aktuell unterstützte Reports:

ABC Analyse

X-Liste (in Vorbereitung)

Die Anwendung ersetzt das frühere AutoHotkey Automationssystem und wird als WPF Desktop Application (.NET) umgesetzt.

Architektur
CopagoAutomation
│
├─ Calibration
│  ├─ CalibrationData
│  ├─ CalibrationModeSet
│  ├─ CalibrationDefinitions
│  ├─ CalibrationProfile
│  ├─ CalibrationPoint
│  ├─ CalibrationProfiles
│  ├─ CalibrationService
│  ├─ CalibrationStorage
│  ├─ CalibrationStepDefinition
│  └─ CalibrationRunner
│
├─ Controls
│  └─ DateRuleSelector
│
├─ Data
│  └─ PosRepository
│
├─ Models
│  └─ AppSettings
│
├─ Services
│  ├─ AutomationService
│  └─ (GEPLANT) PathResolver
│
├─ Automation
│  ├─ AbcAutomation
│  ├─ BoundWindowInfo
│  └─ WindowAutomation
│
├─ ViewModels
│  └─ MainViewModel
│
├─ Windows
│  └─ CalibrationPromptWindow
│
├─ MainWindow.xaml
└─ MainWindow.xaml.cs
Implementierte Systeme
1) Settings System

Speichert Benutzerkonfigurationen.

Speicherort:

AppData\Roaming\CopagoAutomation\settings.json

Gespeicherte Werte:

ABC Mode (Laptop / Docking)

X Mode (Laptop / Docking)

NEU (geplant):

SaveMode (SemcoUpload / Alternativ)

Alternativer Speicherpfad

Verwendete Klassen:

SettingsStore

AppSettings

2) Calibration System

Das Kalibrierungssystem ersetzt das frühere AutoHotkey Mapping System.

Speicherort:

AppData\Roaming\CopagoAutomation\calibration.json
Aktueller Stand

✔ vollständig funktionsfähig
✔ UX optimiert
✔ Schrittbasierte Kalibrierung

Wichtige Änderung dieser Session

❌ Entfernt / fachlich verworfen:

SaveDialogPath als Kalibrierpunkt

👉 Begründung:

Der Speicherpfad wird nicht mehr über UI-Koordinaten gesteuert, sondern:

zentral aus Settings / Logik

per Keyboard in den SaveDialog eingefügt

✔ Beibehalten:

OutputClose (optional für Klick auf "Speichern")

3) Hotkey System

Globale Hotkeys über Win32.

Unterstützt:

Ctrl + Alt + 0–9

NumPad 0–9

Status

✔ stabil
✔ korrekt sequenziell
✔ Anzeige synchron zur Logik

4) Koordinatenberechnung

✔ DPI unabhängig
✔ Multi-Monitor sicher
✔ Root Window basiert

5) Calibration Workflow
Verbesserung

✔ MainWindow wird während Kalibrierung ausgeblendet

👉 Ergebnis:

keine UI-Überlagerung

saubere Benutzerführung

6) Automation System

Zentrale Klasse:

AutomationService

Status

✔ stabil
✔ unterstützt Mode-Handling
✔ saubere Übergabe an Engines

7) Window Automation

✔ Window Binding aktiv
✔ Handle-basierte Steuerung
✔ Guard System integriert

8) Automation Guard System

✔ verhindert Fehlklicks
✔ erkennt Fokusverlust
✔ bricht sicher ab

9) Copago Window Detection

✔ Titelbasierte Erkennung
✔ flexibel gegenüber Fenstertiteln

🔥 NEUE KERNARCHITEKTUR: SAVE-STRATEGIE
Problemstellung (alt)

Bisher:

SaveDialog sollte über Kalibrierpunkte gesteuert werden

Speicherpfad war nicht klar definiert

BaseFolder nicht sauber integriert

👉 Ergebnis:

instabil

unvollständig

nicht produktionsfähig

Neue Lösung (aktuell gültig)
Einführung eines Speicher-Modus-Systems
🔹 Modus 1: Semco Upload (Standard)

Standardbetrieb der Anwendung.

Eigenschaften:

Für jeden:

Bericht

POS

existiert ein fest definierter Zielpfad

👉 Pfade werden nicht vom Benutzer gewählt

Verhalten:

vollständig automatisiert

reproduzierbar

kein UI-Eingriff notwendig

UI:

Speicherpfad-Feld: deaktiviert

Ordner-Auswahl: deaktiviert

🔹 Modus 2: Alternativ

Manueller Modus für Sonderfälle.

Eigenschaften:

Benutzer kann eigenen Speicherpfad wählen

UI:

Speicherpfad-Feld: aktiv

Ordner-Auswahl: aktiv

🧠 Zentrale Logik

Die Anwendung entscheidet zur Laufzeit:

Wenn SaveMode = SemcoUpload:
    → nutze festen Zielpfad

Wenn SaveMode = Alternativ:
    → nutze Benutzerpfad
🔧 Geplante Komponente

Neue zentrale Klasse:

PathResolver (oder SemcoPathResolver)
Aufgabe:
Input:
- Report (z. B. ABC)
- POS
- Modus

Output:
- finaler Speicherpfad
Beispiel (Zielstruktur):
BaseFolder\POS\ABC_Analyse.pdf
10) ABC Automation Workflow
Aktueller Ablauf

POS auswählen

Datum setzen

Report starten

SaveDialog öffnen

Status

✔ Automation startet
✔ Kalibrierpunkte werden genutzt
✔ Window Binding funktioniert

⚠️ Aktueller Stand

SaveDialog ist noch nicht automatisiert

🚧 Aktuelle Blocker
1) SaveDialog Automation

Noch nicht implementiert:

Pfad setzen

Dateiname setzen

Speichern auslösen

2) PathResolver fehlt

keine zentrale Pfadlogik vorhanden

Semco-Modus noch nicht umgesetzt

3) UI-Integration fehlt

Modus-Auswahl nicht vorhanden

Pfad-Feld noch nicht dynamisch steuerbar

📌 Nächste Entwicklungsschritte
🔥 PRIORITÄT 1
SaveMode System implementieren

SemcoUpload / Alternativ

Speicherung in AppSettings

UI-Steuerung (aktiv / deaktiviert)

🔥 PRIORITÄT 2
PathResolver entwickeln

feste Pfade pro:

POS

Report

zentrale Logik

🔥 PRIORITÄT 3
SaveDialog Automation

Ziel:

1. Dialog erkennen
2. Fokus setzen
3. Pfad per Clipboard einfügen
4. Enter
5. Dateiname setzen
6. Enter

👉 ohne Klick-Logik, nur Keyboard

🔥 PRIORITÄT 4
BaseFolder Integration

Grundlage für alle festen Pfade

🔥 PRIORITÄT 5
X-Liste Automation

eigene Engine oder Erweiterung

🔥 PRIORITÄT 6
Stability Upgrade

Retry Mechanik

Re-Focus

Logging

🔥 PRIORITÄT 7
Calibration Overview UI

Punkte anzeigen

bearbeiten

löschen

Teststatus
Bereich	Status
Kalibrierung	✅ OK
Hotkeys	✅ OK
INI / Settings	✅ OK
Window Binding	✅ OK
Automation Start	⚠️ Teilweise
SaveDialog Handling	❌ Offen
PathResolver	❌ Offen
Nächste Session

Start mit:

👉 SaveMode System + UI Integration

Empfohlener nächster Schritt

👉 JETZT:

Speicher-Modus in AppSettings + UI einführen

Ende der PROJECT_STATUS.md