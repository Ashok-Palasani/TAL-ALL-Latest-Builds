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
    
    public partial class tblSPrejectreason
    {
        public int SPRID { get; set; }
        public string RejectName { get; set; }
        public string RejectNameDesc { get; set; }
        public Nullable<int> isDeleted { get; set; }
        public Nullable<System.DateTime> CreatedOn { get; set; }
        public Nullable<int> CreatedBy { get; set; }
        public Nullable<System.DateTime> ModifiedOn { get; set; }
        public Nullable<int> ModifiedBy { get; set; }
        public string CorrectedDate { get; set; }
        public Nullable<int> Cellid { get; set; }
        public Nullable<int> Machineid { get; set; }
        public int IsMaint { get; set; }
        public int IsProd { get; set; }
        public int isTCF { get; set; }
        public int SPHMIID { get; set; }
    }
}
