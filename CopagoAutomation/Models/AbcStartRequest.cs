using System.Collections.Generic;

namespace CopagoAutomation.Models
{
	public class AbcStartRequest
	{
		public MachineMode Mode { get; set; }
            public string BaseFolder { get; set; } = string.Empty;
            public SaveMode SaveMode { get; set; }
            public string? SammelordnerPath { get; set; }
		public bool UseSammelordner { get; set; }
		public int SelectedPosCount { get; set; }

        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        /// <summary>
        /// Wird bei "Wiederkehrend" gesetzt. Jeder Eintrag wird als DateFrom = DateTo verwendet.
        /// Ist null oder leer → einmaliger Zeitraum mit DateFrom/DateTo.
        /// </summary>
        public List<DateTime>? OccurrenceDates { get; set; }
		public List<string> SelectedPosValues { get; set; } = new();
	}
}