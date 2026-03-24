using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
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
                    Sleep(DefaultRunReportWaitMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(outputSavePoint, boundWindow);
                    logs.Add($"Save für POS {currentPos} ausgelöst");
                    Sleep(DefaultSaveDialogWaitMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(outputClosePoint, boundWindow);
                    logs.Add($"Fenster für POS {currentPos} geschlossen");
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
