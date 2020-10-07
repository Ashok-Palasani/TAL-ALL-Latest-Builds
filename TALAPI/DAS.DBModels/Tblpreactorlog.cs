using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblpreactorlog
    {
        public int LogId { get; set; }
        public string OutputFileName { get; set; }
        public string LogInfo { get; set; }
        public DateTime? ReportGenaratedTime { get; set; }
        public int Isgenerated { get; set; }
        public int? ScheduleId { get; set; }
        public string CorrectedDate { get; set; }
    }
}
