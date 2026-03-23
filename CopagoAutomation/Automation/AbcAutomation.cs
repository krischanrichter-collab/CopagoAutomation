using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using CopagoAutomation.Calibration;
using CopagoAutomation.Models;

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

		private const ushort VK_CONTROL = 0x11;
		private const ushort VK_A = 0x41;
		private const ushort VK_RETURN = 0x0D;

		private readonly WindowAutomation _windowAutomation;

		public AbcAutomation()
			: this(new WindowAutomation())
		{
		}

		public AbcAutomation(WindowAutomation windowAutomation)
		{
			_windowAutomation = windowAutomation ?? throw new ArgumentNullException(nameof(windowAutomation));
		}

		public List<string> Run(AbcStartRequest request)
		{
			var logs = new List<string>();
			logs.Add("Fehler: Diese Überladung verwendet noch keine Kalibrierpunkte.");
			logs.Add("Bitte AutomationService so verwenden, dass Run(request, calibrationPoints) aufgerufen wird.");
			return logs;
		}

		public List<string> Run(
			AbcStartRequest request,
			IReadOnlyDictionary<string, CalibrationPoint> calibrationPoints)
		{
			var logs = new List<string>();

			if (request == null)
			{
				logs.Add("Fehler: Request ist null.");
				return logs;
			}

			if (calibrationPoints == null)
			{
				logs.Add("Fehler: Kalibrierpunkte fehlen.");
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

			if (!TryGetRequiredPoint(calibrationPoints, "POS", out var posPoint, logs)) return logs;
			if (!TryGetRequiredPoint(calibrationPoints, "DateFrom", out var dateFromPoint, logs)) return logs;
			if (!TryGetRequiredPoint(calibrationPoints, "DateTo", out var dateToPoint, logs)) return logs;
			if (!TryGetRequiredPoint(calibrationPoints, "RunReport", out var runReportPoint, logs)) return logs;
			if (!TryGetRequiredPoint(calibrationPoints, "OutputSave", out var outputSavePoint, logs)) return logs;
			if (!TryGetRequiredPoint(calibrationPoints, "OutputClose", out var outputClosePoint, logs)) return logs;

			if (!TryResolveDateRange(request, out var dateFrom, out var dateTo, out var dateError))
			{
				logs.Add(dateError);
				return logs;
			}

			if (!_windowAutomation.TryBindWindowByTitleContains(CopagoWindowTitlePart, out var boundWindow))
			{
				logs.Add($"Fehler: Copago Fenster konnte nicht gefunden werden. Erwarteter Titelteil: '{CopagoWindowTitlePart}'.");
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

			string dateFromText = dateFrom.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
			string dateToText = dateTo.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

			logs.Add("ABC Automation gestartet");
			logs.Add($"Gebundenes Fenster: {boundWindow.Title}");
			logs.Add($"Zeitraum: {dateFromText} bis {dateToText}");
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

					ClickPoint(posPoint);
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

					ClickPoint(dateFromPoint);
					Sleep(DefaultActionDelayMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					SelectAll();
					Sleep(80);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					TypeText(dateFromText);
					Sleep(DefaultActionDelayMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					ClickPoint(dateToPoint);
					Sleep(DefaultActionDelayMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					SelectAll();
					Sleep(80);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					TypeText(dateToText);
					Sleep(DefaultActionDelayMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					ClickPoint(runReportPoint);
					logs.Add($"Report für POS {currentPos} gestartet");
					Sleep(DefaultRunReportWaitMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					ClickPoint(outputSavePoint);
					logs.Add($"Save für POS {currentPos} ausgelöst");
					Sleep(DefaultSaveDialogWaitMs);

					if (!EnsureBoundWindowReady(boundWindow, logs))
						return logs;

					ClickPoint(outputClosePoint);
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

		private static bool TryResolveDateRange(
			AbcStartRequest request,
			out DateTime dateFrom,
			out DateTime dateTo,
			out string error)
		{
			dateFrom = default;
			dateTo = default;
			error = string.Empty;

			if (!TryReadDateProperty(request, new[] { "DateFrom", "FromDate", "StartDate", "Von", "From" }, out dateFrom))
			{
				error = "Fehler: Kein gültiges Startdatum im AbcStartRequest gefunden.";
				return false;
			}

			if (!TryReadDateProperty(request, new[] { "DateTo", "ToDate", "EndDate", "Bis", "To" }, out dateTo))
			{
				error = "Fehler: Kein gültiges Enddatum im AbcStartRequest gefunden.";
				return false;
			}

			if (dateTo < dateFrom)
			{
				error = "Fehler: Das Enddatum liegt vor dem Startdatum.";
				return false;
			}

			return true;
		}

		private static bool TryReadDateProperty(
			object instance,
			string[] propertyNames,
			out DateTime value)
		{
			value = default;

			var type = instance.GetType();

			foreach (var propertyName in propertyNames)
			{
				var property = type.GetProperty(
					propertyName,
					BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

				if (property == null)
					continue;

				var rawValue = property.GetValue(instance);

				if (rawValue is DateTime dt)
				{
					value = dt;
					return true;
				}

				if (rawValue is string s &&
					DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
				{
					value = parsed;
					return true;
				}
			}

			return false;
		}

		private static void ClickPoint(CalibrationPoint point)
		{
			SetCursorPos(point.X, point.Y);
			Sleep(60);
			LeftClick();
			Sleep(DefaultClickDelayMs);
		}

		private static void SelectAll()
		{
			KeyDown(VK_CONTROL);
			Sleep(25);
			KeyPress(VK_A);
			Sleep(25);
			KeyUp(VK_CONTROL);
			Sleep(60);
		}

		private static void PressKey(ushort virtualKey)
		{
			KeyPress(virtualKey);
			Sleep(60);
		}

		private static void TypeText(string text)
		{
			if (string.IsNullOrEmpty(text))
				return;

			foreach (char ch in text)
			{
				SendUnicodeChar(ch);
				Sleep(DefaultTypingDelayMs);
			}
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
	}
}