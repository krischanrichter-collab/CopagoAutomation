using System;
using System.Collections.Generic;
using CopagoAutomation.Automation;
using CopagoAutomation.Calibration;
using CopagoAutomation.Models;

namespace CopagoAutomation.Services
{
	public class AutomationService
	{
		private const string AbcProfileName = "ABC Analyse";

		private readonly CalibrationService _calibrationService;
		private readonly AbcAutomation _abcAutomation;

		public AutomationService(CalibrationService calibrationService)
		{
			_calibrationService = calibrationService
				?? throw new ArgumentNullException(nameof(calibrationService));

			_abcAutomation = new AbcAutomation();
		}

		public List<string> StartAbcAutomation(AbcStartRequest request)
		{
			if (request == null)
				throw new ArgumentNullException(nameof(request));

			string modeName = ResolveModeName(request);

			if (!_calibrationService.IsProfileComplete(modeName, AbcProfileName))
			{
				throw new InvalidOperationException(
					$"Die Kalibrierung für Profil '{AbcProfileName}' im Modus '{modeName}' ist nicht vollständig.");
			}

			var requiredPoints = GetRequiredAbcPoints(modeName);

			return _abcAutomation.Run(request, requiredPoints);
		}

		private Dictionary<string, CalibrationPoint> GetRequiredAbcPoints(string modeName)
		{
			var points = new Dictionary<string, CalibrationPoint>(StringComparer.OrdinalIgnoreCase);

			points["POS"] = GetRequiredPoint(modeName, AbcProfileName, "POS");
			points["DateFrom"] = GetRequiredPoint(modeName, AbcProfileName, "DateFrom");
			points["DateTo"] = GetRequiredPoint(modeName, AbcProfileName, "DateTo");
			points["RunReport"] = GetRequiredPoint(modeName, AbcProfileName, "RunReport");
			points["OutputSave"] = GetRequiredPoint(modeName, AbcProfileName, "OutputSave");
			points["OutputClose"] = GetRequiredPoint(modeName, AbcProfileName, "OutputClose");

			return points;
		}

		private CalibrationPoint GetRequiredPoint(string modeName, string profileName, string key)
		{
			var point = _calibrationService.GetPoint(modeName, profileName, key);

			if (point == null)
			{
				throw new InvalidOperationException(
					$"Kalibrierpunkt '{key}' fehlt für Profil '{profileName}' im Modus '{modeName}'.");
			}

			return point;
		}

		private static string ResolveModeName(AbcStartRequest request)
		{
			var type = request.GetType();

			var modeProperty =
				type.GetProperty("Mode")
				?? type.GetProperty("MachineMode")
				?? type.GetProperty("ModeName");

			if (modeProperty == null)
			{
				throw new InvalidOperationException(
					"Im AbcStartRequest wurde keine Modus-Eigenschaft gefunden. Erwartet wird z. B. Mode, MachineMode oder ModeName.");
			}

			object? rawValue = modeProperty.GetValue(request);

			if (rawValue == null)
			{
				throw new InvalidOperationException(
					"Im AbcStartRequest wurde kein gültiger Modus gefunden. Erwartet wird z. B. 'laptop' oder 'dock'.");
			}

			if (rawValue is MachineMode machineMode)
			{
				return machineMode == MachineMode.Laptop ? "laptop" : "dock";
			}

			string rawMode = rawValue.ToString()?.Trim() ?? string.Empty;

			if (string.IsNullOrWhiteSpace(rawMode))
			{
				throw new InvalidOperationException(
					"Im AbcStartRequest wurde kein gültiger Modus gefunden. Erwartet wird z. B. 'laptop' oder 'dock'.");
			}

			string normalized = rawMode.ToLowerInvariant();

			return normalized switch
			{
				"laptop" => "laptop",
				"dock" => "dock",
				"docking" => "dock",
				_ => throw new InvalidOperationException(
					$"Unbekannter Modus '{rawMode}'. Erwartet wird 'laptop' oder 'dock'.")
			};
		}
	}
}