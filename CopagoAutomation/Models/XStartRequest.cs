using System;
using System.Collections.Generic;

namespace CopagoAutomation.Models
{
    public class XStartRequest
    {
        public MachineMode Mode { get; set; }
        public string BaseFolder { get; set; } = string.Empty;
        public bool UseSammelordner { get; set; }
        public int SelectedPosCount { get; set; }
        public List<string> SelectedPosValues { get; set; } = new();

        public int Year { get; set; }
        public int CumPercent { get; set; }
        public int ToWeek { get; set; }

        public SaveMode SaveMode { get; set; }
    }
}
