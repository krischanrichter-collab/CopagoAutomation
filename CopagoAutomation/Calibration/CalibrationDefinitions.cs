using System.Collections.Generic;
using CopagoAutomation.Models;

namespace CopagoAutomation.Calibration
{
	public static class CalibrationDefinitions
	{
		// Schritte die nur im PDF-Modus kalibriert werden (Diskette)
		private static readonly HashSet<string> PdfOnlyKeys = new(System.StringComparer.OrdinalIgnoreCase)
			{ "OutputSave" };

		// Schritte die nur im Excel-Modus kalibriert werden
		private static readonly HashSet<string> ExcelOnlyKeys = new(System.StringComparer.OrdinalIgnoreCase)
			{ "OutputExcelExport", "ConfirmOk" };

		private static bool IsStepApplicable(string key, OutputFormat format) =>
			format == OutputFormat.Excel ? !PdfOnlyKeys.Contains(key) : !ExcelOnlyKeys.Contains(key);
		// Deutsche Tastatur obere Reihe: Q W E R T Z U I O P (Schritte 1–10)

		public static readonly IReadOnlyList<CalibrationStepDefinition> AbcAnalyseSteps = new[]
		{
			new CalibrationStepDefinition
			{
				Key             = "POS",
				Title           = "Dropdown \"Filiale\"",
				HotkeyDigit     = 1,
				HotkeyText      = "Strg+Alt+Q",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Filiale\" und drücke dann Strg+Alt+Q."
			},
			new CalibrationStepDefinition
			{
				Key             = "DateFrom",
				Title           = "Datumauswahl \"Datum von\"",
				HotkeyDigit     = 2,
				HotkeyText      = "Strg+Alt+W",
				InstructionText = "Platziere die Maus mittig auf der Datumauswahl \"Datum von\" und drücke dann Strg+Alt+W."
			},
			new CalibrationStepDefinition
			{
				Key             = "DateTo",
				Title           = "Datumauswahl \"Datum bis\"",
				HotkeyDigit     = 3,
				HotkeyText      = "Strg+Alt+E",
				InstructionText = "Platziere die Maus mittig auf der Datumauswahl \"Datum bis\" und drücke dann Strg+Alt+E."
			},
			new CalibrationStepDefinition
			{
				Key             = "RunReport",
				Title           = "\"Start\" (Report ausführen)",
				HotkeyDigit     = 4,
				HotkeyText      = "Strg+Alt+R",
				InstructionText = "Platziere die Maus mittig auf dem \"Start\" Feld und drücke dann Strg+Alt+R."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputExcelExport",
				Title           = "Excel Export Button (nur Excel-Modus)",
				HotkeyDigit     = 5,
				HotkeyText      = "Strg+Alt+T",
				InstructionText = "Nur für Excel-Modus: Sobald das Ausgabe-Fenster erscheint, platziere die Maus mittig auf dem 'Excel Export' Button und drücke dann Strg+Alt+T.",
				IsRequired      = false
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputSave",
				Title           = "Disketten-Symbol (Speichern)",
				HotkeyDigit     = 5,
				HotkeyText      = "Strg+Alt+T",
				InstructionText = "Platziere die Maus mittig auf dem Disketten-Symbol oben links im Ausgabe-Fenster und drücke dann Strg+Alt+T."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogPath",
				Title           = "Adressleiste (Save-Dialog)",
				HotkeyDigit     = 6,
				HotkeyText      = "Strg+Alt+Z",
				InstructionText = "Platziere die Maus mittig auf der Adressleiste (Pfadanzeige oben) im Save-Dialog und drücke dann Strg+Alt+Z."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogFilename",
				Title           = "Dateiname-Feld (Save-Dialog)",
				HotkeyDigit     = 7,
				HotkeyText      = "Strg+Alt+U",
				InstructionText = "Platziere die Maus mittig auf dem Dateiname-Eingabefeld im Save-Dialog und drücke dann Strg+Alt+U."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputClose",
				Title           = "Speichern-Button (Save-Dialog)",
				HotkeyDigit     = 8,
				HotkeyText      = "Strg+Alt+I",
				InstructionText = "Platziere die Maus mittig auf dem Speichern-Button im Save-Dialog und drücke dann Strg+Alt+I."
			},
			new CalibrationStepDefinition
			{
				Key             = "ConfirmOk",
				Title           = "OK-Button (Bestätigungsmeldung, nur Excel-Modus)",
				HotkeyDigit     = 9,
				HotkeyText      = "Strg+Alt+O",
				InstructionText = "Nur für Excel-Modus: Nach dem Speichern erscheint eine Bestätigungsmeldung. Platziere die Maus auf dem OK-Button und drücke dann Strg+Alt+O.",
				IsRequired      = false
			}
		};

		public static readonly IReadOnlyList<CalibrationStepDefinition> XListeSteps = new[]
		{
			new CalibrationStepDefinition
			{
				Key             = "POS",
				Title           = "Dropdown \"Filiale\"",
				HotkeyDigit     = 1,
				HotkeyText      = "Strg+Alt+Q",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Filiale\" und drücke dann Strg+Alt+Q."
			},
			new CalibrationStepDefinition
			{
				Key             = "Year",
				Title           = "Dropdown \"Jahr\"",
				HotkeyDigit     = 2,
				HotkeyText      = "Strg+Alt+W",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Jahr\" und drücke dann Strg+Alt+W."
			},
			new CalibrationStepDefinition
			{
				Key             = "KwFrom",
				Title           = "Feld \"Von Kalenderwoche\"",
				HotkeyDigit     = 3,
				HotkeyText      = "Strg+Alt+E",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Von Kalenderwoche\" und drücke dann Strg+Alt+E."
			},
			new CalibrationStepDefinition
			{
				Key             = "KwTo",
				Title           = "Feld \"Bis Kalenderwoche\"",
				HotkeyDigit     = 4,
				HotkeyText      = "Strg+Alt+R",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Bis Kalenderwoche\" und drücke dann Strg+Alt+R."
			},
			new CalibrationStepDefinition
			{
				Key             = "Kumul",
				Title           = "Feld \"Kumulierter Anteil bis\"",
				HotkeyDigit     = 5,
				HotkeyText      = "Strg+Alt+T",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Kumulierter Anteil bis\" und drücke dann Strg+Alt+T."
			},
			new CalibrationStepDefinition
			{
				Key             = "RunReport",
				Title           = "\"Start\" (Report ausführen)",
				HotkeyDigit     = 6,
				HotkeyText      = "Strg+Alt+Z",
				InstructionText = "Platziere die Maus mittig auf dem \"Start\" Feld und drücke dann Strg+Alt+Z."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputExcelExport",
				Title           = "Excel Export Button (nur Excel-Modus)",
				HotkeyDigit     = 7,
				HotkeyText      = "Strg+Alt+U",
				InstructionText = "Nur für Excel-Modus: Sobald das Ausgabe-Fenster erscheint, platziere die Maus mittig auf dem 'Excel Export' Button und drücke dann Strg+Alt+U.",
				IsRequired      = false
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputSave",
				Title           = "Disketten-Symbol (Speichern)",
				HotkeyDigit     = 7,
				HotkeyText      = "Strg+Alt+U",
				InstructionText = "Platziere die Maus mittig auf dem Disketten-Symbol oben links im Ausgabe-Fenster und drücke dann Strg+Alt+U."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogPath",
				Title           = "Adressleiste (Save-Dialog)",
				HotkeyDigit     = 8,
				HotkeyText      = "Strg+Alt+I",
				InstructionText = "Platziere die Maus mittig auf der Adressleiste (Pfadanzeige oben) im Save-Dialog und drücke dann Strg+Alt+I."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogFilename",
				Title           = "Dateiname-Feld (Save-Dialog)",
				HotkeyDigit     = 9,
				HotkeyText      = "Strg+Alt+O",
				InstructionText = "Platziere die Maus mittig auf dem Dateiname-Eingabefeld im Save-Dialog und drücke dann Strg+Alt+O."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputClose",
				Title           = "Speichern-Button (Save-Dialog)",
				HotkeyDigit     = 10,
				HotkeyText      = "Strg+Alt+P",
				InstructionText = "Platziere die Maus mittig auf dem Speichern-Button im Save-Dialog und drücke dann Strg+Alt+P."
			},
			new CalibrationStepDefinition
			{
				Key             = "ConfirmOk",
				Title           = "OK-Button (Bestätigungsmeldung, nur Excel-Modus)",
				HotkeyDigit     = 10,
				HotkeyText      = "Strg+Alt+P",
				InstructionText = "Nur für Excel-Modus: Nach dem Speichern erscheint eine Bestätigungsmeldung. Platziere die Maus auf dem OK-Button und drücke dann Strg+Alt+P.",
				IsRequired      = false
			}
		};

		public static readonly IReadOnlyList<CalibrationStepDefinition> StundenleistungSteps = new[]
		{
			new CalibrationStepDefinition
			{
				Key             = "POS",
				Title           = "Dropdown \"Filiale\"",
				HotkeyDigit     = 1,
				HotkeyText      = "Strg+Alt+Q",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Filiale\" und drücke dann Strg+Alt+Q."
			},
			new CalibrationStepDefinition
			{
				Key             = "Datum",
				Title           = "Datumsauswahl \"Datum\"",
				HotkeyDigit     = 2,
				HotkeyText      = "Strg+Alt+W",
				InstructionText = "Platziere die Maus mittig auf dem Datumsfeld und drücke dann Strg+Alt+W."
			},
			new CalibrationStepDefinition
			{
				Key             = "RunReport",
				Title           = "\"Start\" (Report ausführen)",
				HotkeyDigit     = 3,
				HotkeyText      = "Strg+Alt+E",
				InstructionText = "Platziere die Maus mittig auf dem \"Start\" Feld und drücke dann Strg+Alt+E."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputExcelExport",
				Title           = "Excel Export Button (nur Excel-Modus)",
				HotkeyDigit     = 4,
				HotkeyText      = "Strg+Alt+R",
				InstructionText = "Nur für Excel-Modus: Sobald das Ausgabe-Fenster erscheint, platziere die Maus mittig auf dem 'Excel Export' Button und drücke dann Strg+Alt+R.",
				IsRequired      = false
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputSave",
				Title           = "Disketten-Symbol (Speichern)",
				HotkeyDigit     = 4,
				HotkeyText      = "Strg+Alt+R",
				InstructionText = "Platziere die Maus mittig auf dem Disketten-Symbol oben links im Ausgabe-Fenster und drücke dann Strg+Alt+R."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogPath",
				Title           = "Adressleiste (Save-Dialog)",
				HotkeyDigit     = 5,
				HotkeyText      = "Strg+Alt+T",
				InstructionText = "Platziere die Maus mittig auf der Adressleiste (Pfadanzeige oben) im Save-Dialog und drücke dann Strg+Alt+T."
			},
			new CalibrationStepDefinition
			{
				Key             = "SaveDialogFilename",
				Title           = "Dateiname-Feld (Save-Dialog)",
				HotkeyDigit     = 6,
				HotkeyText      = "Strg+Alt+Z",
				InstructionText = "Platziere die Maus mittig auf dem Dateiname-Eingabefeld im Save-Dialog und drücke dann Strg+Alt+Z."
			},
			new CalibrationStepDefinition
			{
				Key             = "OutputClose",
				Title           = "Speichern-Button (Save-Dialog)",
				HotkeyDigit     = 7,
				HotkeyText      = "Strg+Alt+U",
				InstructionText = "Platziere die Maus mittig auf dem Speichern-Button im Save-Dialog und drücke dann Strg+Alt+U."
			},
			new CalibrationStepDefinition
			{
				Key             = "ConfirmOk",
				Title           = "OK-Button (Bestätigungsmeldung, nur Excel-Modus)",
				HotkeyDigit     = 8,
				HotkeyText      = "Strg+Alt+I",
				InstructionText = "Nur für Excel-Modus: Nach dem Speichern erscheint eine Bestätigungsmeldung. Platziere die Maus auf dem OK-Button und drücke dann Strg+Alt+I.",
				IsRequired      = false
			}
		};

		public static IReadOnlyList<CalibrationStepDefinition> GetStepsForProfile(string profileName)
		{
			return profileName switch
			{
				CalibrationProfiles.AbcAnalyse      => AbcAnalyseSteps,
				CalibrationProfiles.XListe          => XListeSteps,
				CalibrationProfiles.Stundenleistung => StundenleistungSteps,
				_                                   => new CalibrationStepDefinition[0]
			};
		}

		public static IReadOnlyList<string> GetKeysForProfile(string profileName)
		{
			var steps = GetStepsForProfile(profileName);
			var keys = new List<string>(steps.Count);
			foreach (var step in steps)
				keys.Add(step.Key);
			return keys;
		}

		public static IReadOnlyList<string> GetRequiredKeysForProfile(string profileName, OutputFormat format = OutputFormat.Pdf)
		{
			var steps = GetStepsForProfile(profileName);
			var keys = new List<string>(steps.Count);
			foreach (var step in steps)
			{
				if (!IsStepApplicable(step.Key, format)) continue;
				// Im Excel-Modus ist OutputExcelExport pflicht (auch wenn IsRequired=false gesetzt)
				bool required = step.IsRequired || (format == OutputFormat.Excel && ExcelOnlyKeys.Contains(step.Key));
				if (required)
					keys.Add(step.Key);
			}
			return keys;
		}
	}
}
