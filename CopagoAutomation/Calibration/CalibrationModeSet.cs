using System.Collections.Generic;

namespace CopagoAutomation.Calibration
{
	public class CalibrationModeSet
	{
		public string ModeName { get; set; } = string.Empty;

		public List<CalibrationProfile> Profiles { get; set; } = new();
	}
}