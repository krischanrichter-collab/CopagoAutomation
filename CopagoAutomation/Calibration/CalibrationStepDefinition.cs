namespace CopagoAutomation.Calibration
{
	public class CalibrationStepDefinition
	{
		public string Key { get; set; } = string.Empty;
		public string Title { get; set; } = string.Empty;
		public string HotkeyText { get; set; } = string.Empty;
		public string InstructionText { get; set; } = string.Empty;

		/// <summary>The digit (1–9) the user must press together with Ctrl+Alt to capture this point.</summary>
		public int HotkeyDigit { get; set; }
	}
}