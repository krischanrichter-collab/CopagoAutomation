using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using CopagoAutomation.Calibration;
using CopagoAutomation.Models;
using CopagoAutomation.Services;

namespace CopagoAutomation.Automation
{
    public class AbcAutomation
    {
        private const int DefaultClickDelayMs = 180;
        private const int DefaultActionDelayMs = 250;
        private const int DefaultTypingDelayMs = 35;
        private const int DefaultRunReportWaitMs = 1500;
        private const int DefaultSaveDialogWaitMs = 700;
        private const int ReportReadyTimeoutMs = 60_000;
        private const int ReportReadyPollIntervalMs = 500;

        private const string CopagoWindowTitlePart = "copago Office Online Verwaltung";

        private readonly PathResolver _pathResolver;
        private readonly WindowAutomation _windowAutomation;
        private readonly CalibrationService _calibrationService;

        public AbcAutomation(PathResolver pathResolver, CalibrationService calibrationService)
            : this(new WindowAutomation(), pathResolver, calibrationService)
        {
        }

        public AbcAutomation(WindowAutomation windowAutomation, PathResolver pathResolver, CalibrationService calibrationService)
        {
            _windowAutomation = windowAutomation ?? throw new ArgumentNullException(nameof(windowAutomation));
            _pathResolver = pathResolver ?? throw new ArgumentNullException(nameof(pathResolver));
            _calibrationService = calibrationService ?? throw new ArgumentNullException(nameof(calibrationService));
        }

        public List<string> Run(
            AbcStartRequest request,
            string calibrationModeName,
            string calibrationProfileName)
        {
            var logs = new List<string>();

            if (request == null)
            {
                logs.Add("Fehler: Request ist null.");
                return logs;
            }

            if (string.IsNullOrWhiteSpace(request.BaseFolder))
            {
                logs.Add("Hinweis: Basisordner ist leer (wird aktuell noch nicht verwendet).");
            }

            if (request.SelectedPosValues == null || !request.SelectedPosValues.Any())
            {
                logs.Add("Fehler: Keine POS ausgewählt.");
                return logs;
            }

            if (!_windowAutomation.TryBindWindowByTitleContains(CopagoWindowTitlePart, out var boundWindow))
            {
                logs.Add($"Fehler: Copago Fenster konnte nicht gefunden werden. Erwarteter Titelteil: \'{CopagoWindowTitlePart}\'");
                return logs;
            }

            if (!_windowAutomation.TryActivateBoundWindow(boundWindow))
            {
                logs.Add("Fehler: Gebundenes Copago Fenster konnte nicht aktiviert werden.");
                return logs;
            }

            Sleep(DefaultActionDelayMs);

            if (!EnsureBoundWindowReady(boundWindow, logs))
                return logs;

            string dateFromText = request.DateFrom.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
            string dateToText = request.DateTo.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

            logs.Add("ABC Automation gestartet");
            logs.Add($"Gebundenes Fenster: {boundWindow.Title}");
            logs.Add($"Zeitraum: {dateFromText} bis {dateToText}");
            logs.Add($"POS Anzahl: {request.SelectedPosValues.Count()}");
            logs.Add("Hinweis: BaseFolder ist noch nicht aktiv in die Save-Dialog-Steuerung eingebunden.");

            var posPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "POS", boundWindow);
            if (posPoint == null) { logs.Add("Fehler: Kalibrierpunkt \'POS\' fehlt."); return logs; }
            var dateFromPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "DateFrom", boundWindow);
            if (dateFromPoint == null) { logs.Add("Fehler: Kalibrierpunkt \'DateFrom\' fehlt."); return logs; }
            var dateToPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "DateTo", boundWindow);
            if (dateToPoint == null) { logs.Add("Fehler: Kalibrierpunkt \'DateTo\' fehlt."); return logs; }
            var runReportPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "RunReport", boundWindow);
            if (runReportPoint == null) { logs.Add("Fehler: Kalibrierpunkt \'RunReport\' fehlt."); return logs; }
            var outputSavePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputSave", boundWindow);
            if (outputSavePoint == null) { logs.Add("Fehler: Kalibrierpunkt \'OutputSave\' fehlt."); return logs; }
            var saveDialogPathPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "SaveDialogPath", boundWindow);
            if (saveDialogPathPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'SaveDialogPath' fehlt."); return logs; }
            var outputClosePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputClose", boundWindow);
            if (outputClosePoint == null) { logs.Add("Fehler: Kalibrierpunkt \'OutputClose\' fehlt."); return logs; }

            foreach (var pos in request.SelectedPosValues)
            {
                if (string.IsNullOrWhiteSpace(pos))
                {
                    logs.Add("Leere POS wurde übersprungen.");
                    continue;
                }

                string currentPos = pos.Trim();
                logs.Add($"POS {currentPos} wird verarbeitet");

                try
                {
                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(posPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.TypeText(currentPos);
                    Sleep(120);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.KeyPress(0x0D); // VK_RETURN
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(dateFromPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.TypeText(dateFromText);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(dateToPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    _windowAutomation.TypeText(dateToText);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(runReportPoint, boundWindow);
                    logs.Add($"Report für POS {currentPos} gestartet");

                    if (!WaitForReportReady(boundWindow, logs))
                        return logs;

                    var windowsBeforeSaveDialog = _windowAutomation.GetVisibleTopLevelWindowHandles();
                    ClickPoint(outputSavePoint, boundWindow);
                    logs.Add($"Diskette für POS {currentPos} geklickt");

                    if (!WaitForSaveDialog(windowsBeforeSaveDialog, logs, out IntPtr saveDialogHandle))
                        return logs;

                    string reportName = "ABC_Analyse";
                    string filePath = _pathResolver.ResolvePath(reportName, currentPos, request.SaveMode);
                    logs.Add($"Speicherpfad: {filePath}");

                    if (!_windowAutomation.SetSaveDialogPath(saveDialogHandle, filePath, out string setPathLog))
                    {
                        logs.Add($"Fehler beim Setzen des Speicherpfads: {setPathLog}");
                        return logs;
                    }
                    logs.Add(setPathLog);
                    Sleep(DefaultActionDelayMs);

                    ClickPoint(outputClosePoint, boundWindow);
                    logs.Add($"Gespeichert für POS {currentPos}");
                    Sleep(DefaultActionDelayMs);
                }
                catch (Exception ex)
                {
                    logs.Add($"Fehler bei POS {currentPos}: {ex.Message}");
                }
            }

            logs.Add("ABC Automation abgeschlossen");
            return logs;
        }

        private bool WaitForSaveDialog(HashSet<IntPtr> windowsBefore, List<string> logs, out IntPtr dialogHandle, int timeoutMs = 10_000)
        {
            dialogHandle = IntPtr.Zero;
            logs.Add("Warte auf Save-Dialog...");
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                var windowsNow = _windowAutomation.GetVisibleTopLevelWindowHandles();
                var newHandle = windowsNow.FirstOrDefault(h => !windowsBefore.Contains(h));
                if (newHandle != IntPtr.Zero)
                {
                    Thread.Sleep(ReportReadyPollIntervalMs);
                    dialogHandle = newHandle;
                    logs.Add("Save-Dialog erkannt.");
                    return true;
                }
                Thread.Sleep(ReportReadyPollIntervalMs);
            }
            logs.Add($"Timeout: Save-Dialog nicht innerhalb von {timeoutMs / 1000}s erschienen.");
            return false;
        }

        private bool EnsureBoundWindowReady(BoundWindowInfo boundWindow, List<string> logs)
        {
            if (!_windowAutomation.IsValidHandle(boundWindow.Handle))
            {
                logs.Add("Automation gestoppt: Gebundenes Copago Fenster ist nicht mehr gültig oder wurde geschlossen.");
                return false;
            }

            // Bereits aktiv → alles gut
            if (_windowAutomation.IsBoundWindowActive(boundWindow))
                return true;

            logs.Add("Hinweis: Copago Fenster ist nicht aktiv. Versuche Re-Activate...");

            const int maxRetries = 2;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                bool activated = _windowAutomation.TryActivateBoundWindow(boundWindow);
                Thread.Sleep(300);

                if (!_windowAutomation.IsValidHandle(boundWindow.Handle))
                {
                    logs.Add("Automation gestoppt: Fenster wurde während Re-Activate geschlossen.");
                    return false;
                }

                if (_windowAutomation.IsBoundWindowActive(boundWindow))
                {
                    logs.Add($"Re-Activate erfolgreich (Versuch {attempt}).");
                    return true;
                }

                logs.Add($"Re-Activate Versuch {attempt} fehlgeschlagen.");
            }

            logs.Add("Automation gestoppt: Fenster konnte nicht reaktiviert werden.");
            return false;
        }

        /// <summary>
        /// Wartet bis das Copago-Fenster wieder auf Nachrichten reagiert (Report fertig ausgewertet).
        /// Startet mit einem kurzen Initialwait damit die App die Auswertung überhaupt beginnen kann,
        /// dann wird alle 500 ms geprüft ob das Fenster wieder reagiert.
        /// </summary>
        /// <summary>
        /// Wartet bis das Report-Output-Fenster erscheint.
        /// Erkennung: Vorher alle sichtbaren Fenster erfassen, dann warten bis ein neues erscheint.
        /// Das ist unabhängig von Fokus oder aktivem Fenster.
        /// </summary>
        private bool WaitForReportReady(BoundWindowInfo boundWindow, List<string> logs)
        {
            logs.Add("Warte auf Fertigstellung des Reports...");

            var windowsBefore = _windowAutomation.GetVisibleTopLevelWindowHandles();
            Thread.Sleep(DefaultRunReportWaitMs); // Initialwait: App startet Auswertung

            var deadline = DateTime.UtcNow.AddMilliseconds(ReportReadyTimeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                if (!_windowAutomation.IsValidHandle(boundWindow.Handle))
                {
                    logs.Add("Fehler: Fenster wurde während der Report-Auswertung geschlossen.");
                    return false;
                }

                var windowsNow = _windowAutomation.GetVisibleTopLevelWindowHandles();
                if (windowsNow.Any(h => !windowsBefore.Contains(h)))
                {
                    Thread.Sleep(ReportReadyPollIntervalMs); // Kurz warten damit das Fenster vollständig gerendert ist
                    logs.Add("Report-Output-Fenster erkannt, Report fertig ausgewertet.");
                    return true;
                }

                Thread.Sleep(ReportReadyPollIntervalMs);
            }

            logs.Add($"Timeout: Report-Output-Fenster nicht innerhalb von {ReportReadyTimeoutMs / 1000}s erschienen.");
            return false;
        }

        private void ClickPoint(CalibrationPoint point, BoundWindowInfo boundWindow)
        {
            // CalibrationService.GetPoint already returns absolute, DPI-scaled coordinates
            _windowAutomation.SetCursorPosition(point.X, point.Y);
            Sleep(60);
            _windowAutomation.LeftClick();
            Sleep(DefaultClickDelayMs);
        }

        private void Sleep(int milliseconds)
        {
            Thread.Sleep(milliseconds);
        }
    }
}
