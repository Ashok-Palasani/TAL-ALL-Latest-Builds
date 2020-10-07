using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblMimicsBackup
    {
        public int Mbid { get; set; }
        public int? Mid { get; set; }
        public string MachineOnTime { get; set; }
        public string OperatingTime { get; set; }
        public string SetupTime { get; set; }
        public string IdleTime { get; set; }
        public string MachineOffTime { get; set; }
        public string BreakdownTime { get; set; }
        public int? MachineId { get; set; }
        public string Shift { get; set; }
        public string CorrectedDate { get; set; }
    }
}
