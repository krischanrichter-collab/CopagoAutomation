using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using CopagoAutomation.Calibration;
using CopagoAutomation.Models;
using CopagoAutomation.Services;

namespace CopagoAutomation.Automation
{
    public class StundenleistungAutomation
    {
        private const int DefaultClickDelayMs       = 180;
        private const int DefaultActionDelayMs      = 250;
        private const int DefaultRunReportWaitMs    = 1500;
        private const int DefaultSaveDialogWaitMs   = 700;
        private const int ReportReadyTimeoutMs      = 60_000;
        private const int ReportReadyPollIntervalMs = 500;

        private const string CopagoWindowTitlePart = "copago Office Online Verwaltung";

        private readonly WindowAutomation    _windowAutomation;
        private readonly PathResolver        _pathResolver;
        private readonly CalibrationService  _calibrationService;

        public StundenleistungAutomation(PathResolver pathResolver, CalibrationService calibrationService)
            : this(new WindowAutomation(), pathResolver, calibrationService) { }

        public StundenleistungAutomation(WindowAutomation windowAutomation, PathResolver pathResolver, CalibrationService calibrationService)
        {
            _windowAutomation   = windowAutomation  ?? throw new ArgumentNullException(nameof(windowAutomation));
            _pathResolver       = pathResolver       ?? throw new ArgumentNullException(nameof(pathResolver));
            _calibrationService = calibrationService ?? throw new ArgumentNullException(nameof(calibrationService));
        }

        public List<string> Run(
            StundenleistungStartRequest request,
            string calibrationModeName,
            string calibrationProfileName,
            CancellationToken ct = default)
        {
            var logs = new List<string>();

            if (request == null)
            {
                logs.Add("Fehler: Request ist null.");
                return logs;
            }

            if (request.SelectedPosValues == null || !request.SelectedPosValues.Any())
            {
                logs.Add("Fehler: Keine POS ausgewählt.");
                return logs;
            }

            if (request.Dates == null || !request.Dates.Any())
            {
                logs.Add("Fehler: Kein Datum ausgewählt.");
                return logs;
            }

            if (!_windowAutomation.TryBindWindowByTitleContains(CopagoWindowTitlePart, out var boundWindow))
            {
                logs.Add($"Fehler: Copago Fenster konnte nicht gefunden werden. Erwarteter Titelteil: '{CopagoWindowTitlePart}'");
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

            var posPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "POS", boundWindow);
            if (posPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'POS' fehlt."); return logs; }
            var datumPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "Datum", boundWindow);
            if (datumPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'Datum' fehlt."); return logs; }
            var runReportPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "RunReport", boundWindow);
            if (runReportPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'RunReport' fehlt."); return logs; }
            var outputSavePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputSave", boundWindow);
            if (outputSavePoint == null) { logs.Add("Fehler: Kalibrierpunkt 'OutputSave' fehlt."); return logs; }
            var saveDialogPathPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "SaveDialogPath", boundWindow);
            if (saveDialogPathPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'SaveDialogPath' fehlt."); return logs; }
            var saveDialogFilenamePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "SaveDialogFilename", boundWindow);
            if (saveDialogFilenamePoint == null) { logs.Add("Fehler: Kalibrierpunkt 'SaveDialogFilename' fehlt."); return logs; }
            var outputClosePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputClose", boundWindow);
            if (outputClosePoint == null) { logs.Add("Fehler: Kalibrierpunkt 'OutputClose' fehlt."); return logs; }

            var outputExcelExportPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputExcelExport", boundWindow);
            bool isExcel = request.OutputFormat == OutputFormat.Excel;
            if (isExcel && outputExcelExportPoint == null)
            {
                logs.Add("Fehler: Kalibrierpunkt 'OutputExcelExport' fehlt (Excel-Modus).");
                return logs;
            }
            string extension = isExcel ? ".xlsx" : ".pdf";

            logs.Add("Stundenleistung Automation gestartet");
            logs.Add($"Gebundenes Fenster: {boundWindow.Title}");
            logs.Add($"Datumseinträge: {request.Dates.Count}");
            logs.Add($"POS Anzahl: {request.SelectedPosValues.Count}");

            foreach (var pos in request.SelectedPosValues)
            {
                if (string.IsNullOrWhiteSpace(pos))
                {
                    logs.Add("Leere POS wurde übersprungen.");
                    continue;
                }

                string currentPos = pos.Trim();

                foreach (var date in request.Dates)
                {
                    string currentDateText = date.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                    ct.ThrowIfCancellationRequested();
                    logs.Add($"POS {currentPos} / Datum {currentDateText} wird verarbeitet");

                    try
                    {
                        if (!EnsureBoundWindowReady(boundWindow, logs))
                            return logs;

                        // Filiale (POS) eintragen
                        ClickPoint(posPoint, boundWindow);
                        Sleep(DefaultActionDelayMs);

                        if (!EnsureBoundWindowReady(boundWindow, logs))
                            return logs;

                        _windowAutomation.SelectAll();
                        Sleep(80);
                        _windowAutomation.TypeText(currentPos);
                        Sleep(120);
                        _windowAutomation.KeyPress(0x0D); // VK_RETURN
                        Sleep(DefaultActionDelayMs);

                        if (!EnsureBoundWindowReady(boundWindow, logs))
                            return logs;

                        // Datum eintragen
                        ClickPoint(datumPoint, boundWindow);
                        Sleep(DefaultActionDelayMs);

                        if (!EnsureBoundWindowReady(boundWindow, logs))
                            return logs;

                        _windowAutomation.SelectAll();
                        Sleep(80);
                        _windowAutomation.TypeText(currentDateText);
                        Sleep(DefaultActionDelayMs);

                        if (!EnsureBoundWindowReady(boundWindow, logs))
                            return logs;

                        // Report starten
                        ClickPoint(runReportPoint, boundWindow);
                        logs.Add($"Report für POS {currentPos} / {currentDateText} gestartet");

                        if (!WaitForReportReady(boundWindow, logs, out IntPtr outputWindowHandle, ct))
                            return logs;

                        // Speichern
                        _windowAutomation.TryActivateWindow(outputWindowHandle);
                        Sleep(DefaultActionDelayMs);

                        var windowsBeforeSaveDialog = _windowAutomation.GetVisibleTopLevelWindowHandles();
                        var saveClickPoint = isExcel ? outputExcelExportPoint! : outputSavePoint;
                        ClickPoint(saveClickPoint, boundWindow);
                        logs.Add(isExcel ? $"Excel Export für POS {currentPos} geklickt" : $"Diskette für POS {currentPos} geklickt");

                        string dateLabel = date.ToString("dd.MM.yy", CultureInfo.InvariantCulture);
                        string filePath  = _pathResolver.ResolvePath("Stundenleistung", currentPos, request.SaveMode, dateLabel, extension);
                        string fileDir   = Path.GetDirectoryName(filePath) ?? string.Empty;
                        string fileName  = Path.GetFileName(filePath);
                        logs.Add($"Versuche, in Datei zu speichern: {filePath}");

                        if (!string.IsNullOrEmpty(fileDir))
                            Directory.CreateDirectory(fileDir);

                        if (!WaitForSaveDialog(windowsBeforeSaveDialog, logs, out IntPtr saveDialogHandle, ct: ct))
                            return logs;

                        if (isExcel)
                        {
                            if (!_windowAutomation.SetSaveDialogPath(saveDialogHandle, filePath, out string pathMsg))
                            {
                                logs.Add($"Fehler beim Setzen des Speicherpfads: {pathMsg}");
                                return logs;
                            }
                            logs.Add(pathMsg);
                            Sleep(DefaultActionDelayMs);
                            _windowAutomation.KeyPress(0x0D); // Enter bestätigt den Speichern-Dialog
                        }
                        else
                        {
                            ClickPoint(saveDialogPathPoint, boundWindow);
                            Sleep(DefaultActionDelayMs);
                            _windowAutomation.SelectAll();
                            Sleep(80);
                            _windowAutomation.TypeText(fileDir);
                            Sleep(DefaultActionDelayMs);
                            _windowAutomation.KeyPress(0x0D);
                            Sleep(DefaultActionDelayMs);

                            ClickPoint(saveDialogFilenamePoint, boundWindow);
                            Sleep(DefaultActionDelayMs);
                            _windowAutomation.SelectAll();
                            Sleep(80);
                            _windowAutomation.TypeText(fileName);
                            Sleep(DefaultActionDelayMs);

                            ClickPoint(outputClosePoint, boundWindow);
                        }

                        logs.Add($"Gespeichert für POS {currentPos} / {currentDateText}");
                        Sleep(DefaultSaveDialogWaitMs);

                        if (isExcel)
                        {
                            logs.Add("Warte auf Excel-Bestätigungsdialog...");
                            var windowsBeforeConfirm = _windowAutomation.GetVisibleTopLevelWindowHandles();
                            _windowAutomation.WaitForNewWindowAndPressEnter(windowsBeforeConfirm, ct: ct);
                            logs.Add("Excel-Bestätigungsdialog bestätigt.");
                        }

                        logs.Add("Warte auf Schließen des Report-Output-Fensters...");
                        _windowAutomation.CloseWindowAndWait(outputWindowHandle, ct: ct);
                        logs.Add("Report-Output-Fenster geschlossen.");
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Fehler bei POS {currentPos} / {currentDateText}: {ex.Message}");
                    }
                }
            }

            logs.Add("Stundenleistung Automation abgeschlossen");
            return logs;
        }

        private bool WaitForSaveDialog(HashSet<IntPtr> windowsBefore, List<string> logs, out IntPtr dialogHandle, int timeoutMs = 10_000, CancellationToken ct = default)
        {
            dialogHandle = IntPtr.Zero;
            logs.Add("Warte auf Save-Dialog...");
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();
                var windowsNow = _windowAutomation.GetVisibleTopLevelWindowHandles();
                var newHandle  = windowsNow.FirstOrDefault(h => !windowsBefore.Contains(h));
                if (newHandle != IntPtr.Zero)
                {
                    ct.WaitHandle.WaitOne(ReportReadyPollIntervalMs);
                    ct.ThrowIfCancellationRequested();
                    dialogHandle = newHandle;
                    logs.Add("Save-Dialog erkannt.");
                    return true;
                }
                ct.WaitHandle.WaitOne(ReportReadyPollIntervalMs);
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

            if (_windowAutomation.IsBoundWindowActive(boundWindow))
                return true;

            logs.Add("Hinweis: Copago Fenster ist nicht aktiv. Versuche Re-Activate...");

            for (int attempt = 1; attempt <= 2; attempt++)
            {
                _windowAutomation.TryActivateBoundWindow(boundWindow);
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

        private bool WaitForReportReady(BoundWindowInfo boundWindow, List<string> logs, out IntPtr outputWindowHandle, CancellationToken ct = default)
        {
            outputWindowHandle = IntPtr.Zero;
            logs.Add("Warte auf Fertigstellung des Reports...");

            var windowsBefore = _windowAutomation.GetVisibleTopLevelWindowHandles();
            ct.WaitHandle.WaitOne(DefaultRunReportWaitMs);
            ct.ThrowIfCancellationRequested();

            var deadline = DateTime.UtcNow.AddMilliseconds(ReportReadyTimeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                ct.ThrowIfCancellationRequested();

                if (!_windowAutomation.IsValidHandle(boundWindow.Handle))
                {
                    logs.Add("Fehler: Fenster wurde während der Report-Auswertung geschlossen.");
                    return false;
                }

                var windowsNow = _windowAutomation.GetVisibleTopLevelWindowHandles();
                var newHandle  = windowsNow.FirstOrDefault(h => !windowsBefore.Contains(h));
                if (newHandle != IntPtr.Zero)
                {
                    ct.WaitHandle.WaitOne(ReportReadyPollIntervalMs);
                    ct.ThrowIfCancellationRequested();
                    outputWindowHandle = newHandle;
                    logs.Add("Report-Output-Fenster erkannt, Report fertig ausgewertet.");
                    return true;
                }

                ct.WaitHandle.WaitOne(ReportReadyPollIntervalMs);
            }

            logs.Add($"Timeout: Report-Output-Fenster nicht innerhalb von {ReportReadyTimeoutMs / 1000}s erschienen.");
            return false;
        }

        private void ClickPoint(CalibrationPoint point, BoundWindowInfo boundWindow)
        {
            _windowAutomation.SetCursorPosition(point.X, point.Y);
            Sleep(60);
            _windowAutomation.LeftClick();
            Sleep(DefaultClickDelayMs);
        }

        private static void Sleep(int milliseconds) => Thread.Sleep(milliseconds);
    }
}
