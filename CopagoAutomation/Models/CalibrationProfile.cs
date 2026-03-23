using System.Collections.Generic;

namespace CopagoAutomation.Models
{
	public sealed class CalibrationProfile
	{
		// z.B. "MAP_Report_Menu", "ExplorerListPoint", "DropPoint", ...
		public Dictionary<string, Point2> Points { get; set; } = new();

		public bool TryGet(string key, out Point2 p) => Points.TryGetValue(key, out p);

		public void Set(string key, int x, int y) => Points[key] = new Point2 { X = x, Y = y };

		public static CalibrationProfile CreateDefault()
		{
			var p = new CalibrationProfile();
			// Defaults optional – erstmal leer/0 ist okay
			p.Set("MAP_Report_Menu", 0, 0);
			p.Set("MAP_Report_Submenu", 0, 0);
			p.Set("ExplorerListPoint", 0, 0);
			p.Set("DropPoint", 0, 0);
			p.Set("RemotePathBox", 0, 0);
			return p;
		}
	}

	public struct Point2
	{
		public int X { get; set; }
		public int Y { get; set; }
	}
}