using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblpreactorschedule
    {
        public int ScheduleId { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public int IsDeleted { get; set; }
        public DateTime InsertedOn { get; set; }
        public int InsertedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string Path { get; set; }
        public int? IsNetwork { get; set; }
        public string DomainName { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public TimeSpan? OutputGenerationTime { get; set; }
        public string FileFormat { get; set; }
        public int IsStart { get; set; }
    }
}
