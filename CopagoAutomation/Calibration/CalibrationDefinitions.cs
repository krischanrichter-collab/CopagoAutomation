using System.Collections.Generic;

namespace CopagoAutomation.Calibration
{
	public static class CalibrationDefinitions
	{
		public static readonly IReadOnlyList<CalibrationStepDefinition> AbcAnalyseSteps = new[]
		{
			new CalibrationStepDefinition
			{
				Key = "POS",
				Title = "Dropdown \"Filiale\"",
				HotkeyText = "Strg+Alt+1",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Filiale\" und drücke dann Strg+Alt+1."
			},
			new CalibrationStepDefinition
			{
				Key = "DateFrom",
				Title = "Datumauswahl \"Datum von\"",
				HotkeyText = "Strg+Alt+2",
				InstructionText = "Platziere die Maus mittig auf der Datumauswahl \"Datum von\" und drücke dann Strg+Alt+2."
			},
			new CalibrationStepDefinition
			{
				Key = "DateTo",
				Title = "Datumauswahl \"Datum bis\"",
				HotkeyText = "Strg+Alt+3",
				InstructionText = "Platziere die Maus mittig auf der Datumauswahl \"Datum bis\" und drücke dann Strg+Alt+3."
			},
			new CalibrationStepDefinition
			{
				Key = "RunReport",
				Title = "\"Start\" (Report ausführen)",
				HotkeyText = "Strg+Alt+4",
				InstructionText = "Platziere die Maus mittig auf dem \"Start\" Feld und führe nach Drücken von Strg+Alt+4 wirklich den Report aus. Dies brauchst Du für die Ausgabe-Schritte."
			},
			new CalibrationStepDefinition
			{
				Key = "OutputSave",
				Title = "Disketten-Symbol (Speichern)",
				HotkeyText = "Strg+Alt+5",
				InstructionText = "Platziere die Maus mittig auf dem Disketten-Symbol oben links im Ausgabe-Fenster und drücke dann Strg+Alt+5."
			},
			new CalibrationStepDefinition
			{
				Key = "SaveDialogPath",
				Title = "Speicherpfad-Feld (Save-Dialog)",
				HotkeyText = "Strg+Alt+6",
				InstructionText = "Platziere die Maus mittig auf dem Feld für den Speicherpfad im Save-Dialog und drücke dann Strg+Alt+6."
			},
			new CalibrationStepDefinition
			{
				Key = "OutputClose",
				Title = "Speichern-Button (Save-Dialog)",
				HotkeyText = "Strg+Alt+7",
				InstructionText = "Platziere die Maus mittig auf dem Speichern-Button im Save-Dialog und drücke dann Strg+Alt+7."
			}
		};

		public static readonly IReadOnlyList<CalibrationStepDefinition> XListeSteps = new[]
		{
			new CalibrationStepDefinition
			{
				Key = "POS",
				Title = "Dropdown \"Filiale\"",
				HotkeyText = "Strg+Alt+1",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Filiale\" und drücke dann Strg+Alt+1."
			},
			new CalibrationStepDefinition
			{
				Key = "Year",
				Title = "Dropdown \"Jahr\"",
				HotkeyText = "Strg+Alt+2",
				InstructionText = "Platziere die Maus mittig auf dem Dropdown \"Jahr\" und drücke dann Strg+Alt+2."
			},
			new CalibrationStepDefinition
			{
				Key = "KwFrom",
				Title = "Feld \"Von Kalenderwoche\"",
				HotkeyText = "Strg+Alt+3",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Von Kalenderwoche\" und drücke dann Strg+Alt+3."
			},
			new CalibrationStepDefinition
			{
				Key = "KwTo",
				Title = "Feld \"Bis Kalenderwoche\"",
				HotkeyText = "Strg+Alt+4",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Bis Kalenderwoche\" und drücke dann Strg+Alt+4."
			},
			new CalibrationStepDefinition
			{
				Key = "Kumul",
				Title = "Feld \"Kumulierter Anteil bis\"",
				HotkeyText = "Strg+Alt+5",
				InstructionText = "Platziere die Maus mittig auf dem Feld \"Kumulierter Anteil bis\" und drücke dann Strg+Alt+5."
			},
			new CalibrationStepDefinition
			{
				Key = "RunReport",
				Title = "\"Start\" (Report ausführen)",
				HotkeyText = "Strg+Alt+6",
				InstructionText = "Platziere die Maus mittig auf dem \"Start\" Feld und führe nach Drücken von Strg+Alt+6 wirklich den Report aus. Dies brauchst Du für die Ausgabe-Schritte."
			},
			new CalibrationStepDefinition
			{
				Key = "OutputSave",
				Title = "Disketten-Symbol (Speichern)",
				HotkeyText = "Strg+Alt+7",
				InstructionText = "Platziere die Maus mittig auf dem Disketten-Symbol oben links im Ausgabe-Fenster und drücke dann Strg+Alt+7."
			},
			new CalibrationStepDefinition
			{
				Key = "SaveDialogPath",
				Title = "Speicherpfad-Feld (Save-Dialog)",
				HotkeyText = "Strg+Alt+8",
				InstructionText = "Platziere die Maus mittig auf dem Feld für den Speicherpfad im Save-Dialog und drücke dann Strg+Alt+8."
			},
			new CalibrationStepDefinition
			{
				Key = "OutputClose",
				Title = "Speichern-Button (Save-Dialog)",
				HotkeyText = "Strg+Alt+9",
				InstructionText = "Platziere die Maus mittig auf dem Speichern-Button im Save-Dialog und drücke dann Strg+Alt+9."
			}
		};

		public static IReadOnlyList<CalibrationStepDefinition> GetStepsForProfile(string profileName)
		{
			return profileName switch
			{
				CalibrationProfiles.AbcAnalyse => AbcAnalyseSteps,
				CalibrationProfiles.XListe => XListeSteps,
				_ => new CalibrationStepDefinition[0]
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
	}
}