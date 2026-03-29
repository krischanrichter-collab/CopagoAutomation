namespace CopagoAutomation.Calibration
{
	public class CalibrationStepDefinition
	{
		public string Key { get; set; } = string.Empty;
		public string Title { get; set; } = string.Empty;
		public string HotkeyText { get; set; } = string.Empty;
		public string InstructionText { get; set; } = string.Empty;

		/// <summary>1-based step index (1–10) mapping to Ctrl+Alt+Q … Ctrl+Alt+P</summary>
		public int HotkeyDigit { get; set; }

		/// <summary>
		/// Pflichtschritt: muss kalibriert sein damit IsProfileComplete true zurückgibt.
		/// Optionale Schritte (z.B. OutputExcelExport) werden nur geprüft wenn der jeweilige Modus aktiv ist.
		/// </summary>
		public bool IsRequired { get; set; } = true;
	}
}