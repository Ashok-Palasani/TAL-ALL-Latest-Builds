//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace MimicsUpdation
{
    using System;
    using System.Collections.Generic;
    
    public partial class tbllossofentry
    {
        public int LossID { get; set; }
        public int MessageCodeID { get; set; }
        public Nullable<System.DateTime> StartDateTime { get; set; }
        public Nullable<System.DateTime> EndDateTime { get; set; }
        public Nullable<System.DateTime> EntryTime { get; set; }
        public string CorrectedDate { get; set; }
        public int MachineID { get; set; }
        public string Shift { get; set; }
        public string MessageDesc { get; set; }
        public string MessageCode { get; set; }
        public int IsUpdate { get; set; }
        public int DoneWithRow { get; set; }
        public Nullable<int> IsStart { get; set; }
        public Nullable<int> IsScreen { get; set; }
        public int ForRefresh { get; set; }
        public Nullable<int> LossMonth { get; set; }
        public Nullable<int> LossYear { get; set; }
        public Nullable<int> LossWeekNumber { get; set; }
        public Nullable<int> LossQuarter { get; set; }
    
        public virtual tbllossescode tbllossescode { get; set; }
    }
}
