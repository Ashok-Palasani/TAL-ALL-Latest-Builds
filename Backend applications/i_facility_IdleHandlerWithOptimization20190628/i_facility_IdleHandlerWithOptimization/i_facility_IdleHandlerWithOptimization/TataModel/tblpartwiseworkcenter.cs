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
    
    public partial class tblpartwiseworkcenter
    {
        public int PartWiseWcId { get; set; }
        public int WorkCenterId { get; set; }
        public int MeasuringUnitId { get; set; }
        public short IsDeleted { get; set; }
        public System.DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public Nullable<System.DateTime> ModifiedOn { get; set; }
        public Nullable<int> ModifiedBy { get; set; }
    
        public virtual tblmachinedetail tblmachinedetail { get; set; }
    }
}
