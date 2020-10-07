using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
   public class PreactorEntity
    {
        public int ScheduleId { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string Path { get; set; }
        public string IsNetwork { get; set; }
        public string DomainName { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string isstart { get;set; }
        public string fileFormat { get; set; }
        public string OutputGenerationTime { get; set; }
    }

    public class ReportEntity
    {
        public DateTime Fromdate { get; set; }
        public DateTime ToDate { get; set; }
        public string Type { get; set; }
    }
}
