namespace CopagoAutomation.Calibration
{
	public class CalibrationPoint
	{
		public string Key { get; set; } = string.Empty;
		public int X { get; set; } // Absolute X-Koordinate
		public int Y { get; set; } // Absolute Y-Koordinate
		public int RelativeX { get; set; } // X-Koordinate relativ zum Fenster
		public int RelativeY { get; set; } // Y-Koordinate relativ zum Fenster
		public bool IsRelative { get; set; } // Gibt an, ob RelativeX/Y verwendet werden sollen
	}
}
