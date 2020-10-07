using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tbltcflossofentry
    {
        public int Ncid { get; set; }
        public int? LossId { get; set; }
        public int? MessageCodeId { get; set; }
        public DateTime? StartDateTime { get; set; }
        public DateTime? EndDateTime { get; set; }
        public string ReasonLevel1 { get; set; }
        public string ReasonLevel2 { get; set; }
        public string ReasonLevel3 { get; set; }
        public string CorrectedDate { get; set; }
        public int? MachineId { get; set; }
        public int IsUpdate { get; set; }
        public int IsArroval { get; set; }
        public int IsAccept { get; set; }
        public int? NoOfReason { get; set; }
        public int? RejectReasonId { get; set; }
        public int ApprovalLevel { get; set; }
        public int IsAccept1 { get; set; }
        public int RejectReasonId1 { get; set; }
        public int UpdateLevel { get; set; }
        public int IsSplitDuration { get; set; }
        public int IsHold { get; set; }
    }
}
