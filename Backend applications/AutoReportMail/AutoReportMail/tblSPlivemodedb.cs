//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AutoReportMail
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblSPlivemodedb
    {
        public int SPModeID { get; set; }
        public int MachineID { get; set; }
        public string Mode { get; set; }
        public System.DateTime InsertedOn { get; set; }
        public int InsertedBy { get; set; }
        public Nullable<System.DateTime> ModifiedOn { get; set; }
        public Nullable<int> ModifiedBy { get; set; }
        public string CorrectedDate { get; set; }
        public int IsDeleted { get; set; }
        public Nullable<System.DateTime> StartTime { get; set; }
        public Nullable<System.DateTime> EndTime { get; set; }
        public string ColorCode { get; set; }
        public int IsCompleted { get; set; }
        public Nullable<int> DurationInSec { get; set; }
        public Nullable<int> IsIdle { get; set; }
        public Nullable<int> IsBreakdown { get; set; }
        public Nullable<int> IsPM { get; set; }
        public Nullable<int> IsGeneric { get; set; }
        public string BatchNumber { get; set; }
        public Nullable<int> Bhmiid { get; set; }
    }
}
