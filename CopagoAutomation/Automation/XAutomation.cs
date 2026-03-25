using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using CopagoAutomation.Calibration;
using CopagoAutomation.Models;
using CopagoAutomation.Services;

namespace CopagoAutomation.Automation
{
    public class XAutomation
    {
        private const int DefaultClickDelayMs = 180;
        private const int DefaultActionDelayMs = 250;
        private const int DefaultTypingDelayMs = 35;
        private const int DefaultRunReportWaitMs = 1500;
        private const int DefaultSaveDialogWaitMs = 700;

        private const string CopagoWindowTitlePart = "copago Office Online Verwaltung";

        private const ushort VK_CONTROL = 0x11;
        private const ushort VK_A = 0x41;
        private const ushort VK_RETURN = 0x0D;

        private readonly WindowAutomation _windowAutomation;
        private readonly PathResolver _pathResolver;
        private readonly CalibrationService _calibrationService;

        public XAutomation(PathResolver pathResolver, CalibrationService calibrationService)
            : this(new WindowAutomation(), pathResolver, calibrationService)
        {
        }

        public XAutomation(WindowAutomation windowAutomation, PathResolver pathResolver, CalibrationService calibrationService)
        {
            _windowAutomation = windowAutomation ?? throw new ArgumentNullException(nameof(windowAutomation));
            _pathResolver = pathResolver ?? throw new ArgumentNullException(nameof(pathResolver));
            _calibrationService = calibrationService ?? throw new ArgumentNullException(nameof(calibrationService));
        }

        public List<string> Run(
            XStartRequest request,
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
                logs.Add($"Fehler: Copago Fenster konnte nicht gefunden werden. Erwarteter Titelteil: '{CopagoWindowTitlePart}'.");
                return logs;
            }

            var posPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "POS", boundWindow);
            if (posPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'POS' fehlt."); return logs; }
            var yearPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "Year", boundWindow);
            if (yearPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'Year' fehlt."); return logs; }
            var cumPercentPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "CumPercent", boundWindow);
            if (cumPercentPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'CumPercent' fehlt."); return logs; }
            var toWeekPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "ToWeek", boundWindow);
            if (toWeekPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'ToWeek' fehlt."); return logs; }
            var runReportPoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "RunReport", boundWindow);
            if (runReportPoint == null) { logs.Add("Fehler: Kalibrierpunkt 'RunReport' fehlt."); return logs; }
            var outputSavePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputSave", boundWindow);
            if (outputSavePoint == null) { logs.Add("Fehler: Kalibrierpunkt 'OutputSave' fehlt."); return logs; }
            var outputClosePoint = _calibrationService.GetPoint(calibrationModeName, calibrationProfileName, "OutputClose", boundWindow);
            if (outputClosePoint == null) { logs.Add("Fehler: Kalibrierpunkt 'OutputClose' fehlt."); return logs; }

            if (!_windowAutomation.TryActivateBoundWindow(boundWindow))
            {
                logs.Add("Fehler: Gebundenes Copago Fenster konnte nicht aktiviert werden.");
                return logs;
            }

            Sleep(DefaultActionDelayMs);

            if (!EnsureBoundWindowReady(boundWindow, logs))
                return logs;

            logs.Add("X-Liste Automation gestartet");
            logs.Add($"Gebundenes Fenster: {boundWindow.Title}");
            logs.Add($"Jahr: {request.Year}");
            logs.Add($"Kumuliert bis %: {request.CumPercent}");
            logs.Add($"Bis KW: {request.ToWeek}");
            logs.Add($"POS Anzahl: {request.SelectedPosValues.Count()}");
            logs.Add("Hinweis: BaseFolder ist noch nicht aktiv in die Save-Dialog-Steuerung eingebunden.");

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

                    SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    TypeText(currentPos);
                    Sleep(120);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    PressKey(VK_RETURN);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(yearPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    TypeText(request.Year.ToString());
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(cumPercentPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    TypeText(request.CumPercent.ToString());
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    ClickPoint(toWeekPoint, boundWindow);
                    Sleep(DefaultActionDelayMs);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    SelectAll();
                    Sleep(80);

                    if (!EnsureBoundWindowReady(boundWindow, logs))
                        return logs;

                    TypeText(request.ToWeek.ToString());
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

                    // Automate Save Dialog
                    string reportName = "X-Liste"; // This should be dynamic if needed
                    string filePath = _pathResolver.ResolvePath(reportName, currentPos, request.SaveMode);
                    logs.Add($"Versuche, in Datei zu speichern: {filePath}");

                    if (!_windowAutomation.AutomateSaveDialog(filePath, "Speichern unter", out string saveDialogLog))
                    {
                        logs.Add($"Fehler bei Save Dialog Automatisierung: {saveDialogLog}");
                        return logs;
                    }
                    logs.Add(saveDialogLog);
                    Sleep(DefaultActionDelayMs); // Give some time after save dialog closes

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

            logs.Add("X-Liste Automation abgeschlossen");
            return logs;
        }

        private bool EnsureBoundWindowReady(BoundWindowInfo boundWindow, List<string> logs)
        {
            if (!_windowAutomation.IsWindowValid(boundWindow))
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

                if (!_windowAutomation.IsWindowValid(boundWindow))
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

        private static bool TryGetRequiredPoint(
            IReadOnlyDictionary<string, CalibrationPoint> calibrationPoints,
            string key,
            out CalibrationPoint point,
            List<string> logs)
        {
            if (calibrationPoints.TryGetValue(key, out point!) && point != null)
                return true;

            logs.Add($"Fehler: Kalibrierpunkt '{key}' fehlt.");
            point = null!;
            return false;
        }

        private void ClickPoint(CalibrationPoint point, BoundWindowInfo boundWindow)
        {
            _windowAutomation.SetCursorPosition(point.X, point.Y);
            Sleep(60);
            _windowAutomation.LeftClick();
            Sleep(DefaultClickDelayMs);
        }

        private void SelectAll()
        {
            _windowAutomation.SelectAll();
            Sleep(60);
        }

        private void PressKey(ushort virtualKey)
        {
            _windowAutomation.KeyPress(virtualKey);
            Sleep(60);
        }

        private void TypeText(string text)
        {
            _windowAutomation.TypeText(text);
            Sleep(DefaultTypingDelayMs);
        }

        private static void Sleep(int milliseconds)
        {
            Thread.Sleep(milliseconds);
        }

        private static void LeftClick()
        {
            var inputs = new INPUT[2];

            inputs[0] = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MOUSEEVENTF_LEFTDOWN
                    }
                }
            };

            inputs[1] = new INPUT
            {
                type = INPUT_MOUSE,
                U = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dwFlags = MOUSEEVENTF_LEFTUP
                    }
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        private static void KeyPress(ushort virtualKey)
        {
            KeyDown(virtualKey);
            Sleep(20);
            KeyUp(virtualKey);
        }

        private static void KeyDown(ushort virtualKey)
        {
            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
                        dwFlags = 0
                    }
                }
            };

            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
        }

        private static void KeyUp(ushort virtualKey)
        {
            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
                        dwFlags = KEYEVENTF_KEYUP
                    }
                }
            };

            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
        }

        private static void SendUnicodeChar(char ch)
        {
            var keyDown = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE
                    }
                }
            };

            var keyUp = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
                    }
                }
            };

            var inputs = new[] { keyDown, keyUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        private const int INPUT_MOUSE = 0;
        private const int INPUT_KEYBOARD = 1;

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public int type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }
    }
}
