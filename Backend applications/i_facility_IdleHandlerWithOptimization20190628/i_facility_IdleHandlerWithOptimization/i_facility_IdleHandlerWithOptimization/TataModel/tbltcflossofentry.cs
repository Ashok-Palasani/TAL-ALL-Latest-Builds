//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace i_facility_IdleHandlerWithOptimization.TataModel
{
    using System;
    using System.Collections.Generic;
    
    public partial class tbltcflossofentry
    {
        public int Ncid { get; set; }
        public Nullable<int> LossId { get; set; }
        public Nullable<int> MessageCodeId { get; set; }
        public Nullable<System.DateTime> StartDateTime { get; set; }
        public Nullable<System.DateTime> EndDateTime { get; set; }
        public string ReasonLevel1 { get; set; }
        public string ReasonLevel2 { get; set; }
        public string ReasonLevel3 { get; set; }
        public string CorrectedDate { get; set; }
        public Nullable<int> MachineId { get; set; }
        public int IsUpdate { get; set; }
        public int IsArroval { get; set; }
        public int IsAccept { get; set; }
        public Nullable<int> NoOfReason { get; set; }
        public Nullable<int> RejectReasonId { get; set; }
        public int ApprovalLevel { get; set; }
        public int IsAccept1 { get; set; }
        public int RejectReasonId1 { get; set; }
        public int UpdateLevel { get; set; }
        public int isSplitDuration { get; set; }
    }
}
