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
        private const string XProfileName = "X-Liste";

		private readonly CalibrationService _calibrationService;
        private readonly AbcAutomation _abcAutomation;
        private readonly XAutomation _xAutomation;

		private readonly PathResolver _pathResolver;

		public AutomationService(CalibrationService calibrationService, PathResolver pathResolver)
		{
			_calibrationService = calibrationService
				?? throw new ArgumentNullException(nameof(calibrationService));
			_pathResolver = pathResolver
				?? throw new ArgumentNullException(nameof(pathResolver));

            _abcAutomation = new AbcAutomation(_pathResolver, _calibrationService);
            _xAutomation = new XAutomation(_pathResolver, _calibrationService);
		}

        public List<string> StartAbcAutomation(AbcStartRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            string modeName = ResolveModeName(request.Mode);

            if (!_calibrationService.IsProfileComplete(modeName, AbcProfileName))
            {
                throw new InvalidOperationException(
                    $"Die Kalibrierung für Profil '{AbcProfileName}' im Modus '{modeName}' ist nicht vollständig.");
            }

            var requiredPoints = GetRequiredAbcPoints(modeName);

            return _abcAutomation.Run(request, modeName, AbcProfileName);
        }

        public List<string> StartXAutomation(XStartRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            string modeName = ResolveModeName(request.Mode);

            if (!_calibrationService.IsProfileComplete(modeName, XProfileName))
            {
                throw new InvalidOperationException(
                    $"Die Kalibrierung für Profil '{XProfileName}' im Modus '{modeName}' ist nicht vollständig.");
            }

            var requiredPoints = GetRequiredXPoints(modeName);

            return _xAutomation.Run(request, modeName, XProfileName);
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

        private Dictionary<string, CalibrationPoint> GetRequiredXPoints(string modeName)
        {
            var points = new Dictionary<string, CalibrationPoint>(StringComparer.OrdinalIgnoreCase);

            points["POS"] = GetRequiredPoint(modeName, XProfileName, "POS");
            points["Year"] = GetRequiredPoint(modeName, XProfileName, "Year");
            points["CumPercent"] = GetRequiredPoint(modeName, XProfileName, "CumPercent");
            points["ToWeek"] = GetRequiredPoint(modeName, XProfileName, "ToWeek");
            points["RunReport"] = GetRequiredPoint(modeName, XProfileName, "RunReport");
            points["OutputSave"] = GetRequiredPoint(modeName, XProfileName, "OutputSave");
            points["OutputClose"] = GetRequiredPoint(modeName, XProfileName, "OutputClose");

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

        private static string ResolveModeName(MachineMode mode)
		{
            return mode == MachineMode.Laptop ? "laptop" : "dock";
		}
	}
}
