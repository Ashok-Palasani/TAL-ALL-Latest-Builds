//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace i_facility.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class tbl_PrevOperationCancel
    {
        public int OPCancelID { get; set; }
        public string ProductionOrder { get; set; }
        public string Operation { get; set; }
        public int IsCancelled { get; set; }
        public Nullable<int> Qty { get; set; }
        public string CorrectedDate { get; set; }
        public string WorkCenter { get; set; }
        public Nullable<System.DateTime> CreatedOn { get; set; }
        public Nullable<int> CreatedBy { get; set; }
        public string PartNumber { get; set; }
    }
}
