using System.Collections.Generic;

namespace CopagoAutomation.Calibration
{
	public class CalibrationProfile
	{
		public string ProfileName { get; set; } = string.Empty;
		public List<CalibrationPoint> Points { get; set; } = new();
	}
}