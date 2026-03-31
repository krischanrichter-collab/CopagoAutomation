using System;
using System.Collections.Generic;
using CopagoAutomation.Models;

namespace CopagoAutomation.Calibration
{
	public class CalibrationRunner
	{
		private readonly List<CalibrationStepDefinition> _steps;

		public string ProfileName { get; }

		public int CurrentIndex { get; private set; }

		public bool HasSteps => _steps.Count > 0;

		public bool IsFinished => !HasSteps || CurrentIndex >= _steps.Count;

		public CalibrationStepDefinition? CurrentStep
		{
			get
			{
				if (IsFinished)
					return null;

				return _steps[CurrentIndex];
			}
		}

		public IReadOnlyList<CalibrationStepDefinition> Steps => _steps;

		public CalibrationRunner(string profileName, OutputFormat format = OutputFormat.Pdf)
		{
			if (string.IsNullOrWhiteSpace(profileName))
				throw new ArgumentException("Profile name must not be empty.", nameof(profileName));

			ProfileName = profileName;

			var all = CalibrationDefinitions.GetStepsForProfile(profileName);
			_steps = new List<CalibrationStepDefinition>();
			foreach (var step in all)
			{
				bool applicable = format == OutputFormat.Excel
					? !IsExclusivelPdf(step.Key)
					: !IsExclusivelyExcel(step.Key);
				if (applicable)
					_steps.Add(step);
			}

			CurrentIndex = 0;
		}

		private static bool IsExclusivelPdf(string key) => key is "OutputSave";

		private static bool IsExclusivelyExcel(string key) =>
			key is "OutputExcelExport" or "ConfirmOk";

		public void Reset()
		{
			CurrentIndex = 0;
		}

		public bool MoveNext()
		{
			if (IsFinished)
				return false;

			CurrentIndex++;
			return !IsFinished;
		}
	}
}