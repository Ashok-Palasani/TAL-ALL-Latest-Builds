using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSplivemodedb
    {
        public int SpmodeId { get; set; }
        public int MachineId { get; set; }
        public string Mode { get; set; }
        public DateTime InsertedOn { get; set; }
        public int InsertedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string CorrectedDate { get; set; }
        public int IsDeleted { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string ColorCode { get; set; }
        public int IsCompleted { get; set; }
        public int? DurationInSec { get; set; }
        public int? IsIdle { get; set; }
        public int? IsBreakdown { get; set; }
        public int? IsPm { get; set; }
        public int? IsGeneric { get; set; }
        public string BatchNumber { get; set; }
        public int? Bhmiid { get; set; }
    }
}
