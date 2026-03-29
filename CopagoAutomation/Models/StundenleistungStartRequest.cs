using System;
using System.Collections.Generic;

namespace CopagoAutomation.Models
{
    public class StundenleistungStartRequest
    {
        public MachineMode Mode { get; set; }
        public string BaseFolder { get; set; } = string.Empty;
        public SaveMode SaveMode { get; set; }
        public string? SammelordnerPath { get; set; }
        public List<string> SelectedPosValues { get; set; } = new();
        public List<DateTime> Dates { get; set; } = new();
        public OutputFormat OutputFormat { get; set; } = OutputFormat.Pdf;
    }
}
