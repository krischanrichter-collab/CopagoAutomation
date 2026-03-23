using System.Collections.Generic;

namespace CopagoAutomation.Models
{
	public class AbcStartRequest
	{
		public MachineMode Mode { get; set; }
		public string BaseFolder { get; set; } = string.Empty;
		public bool UseSammelordner { get; set; }
		public int SelectedPosCount { get; set; }

        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
		public List<string> SelectedPosValues { get; set; } = new();
	}
}