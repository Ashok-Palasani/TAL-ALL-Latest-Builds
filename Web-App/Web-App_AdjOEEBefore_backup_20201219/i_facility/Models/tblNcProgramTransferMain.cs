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
    
    public partial class tblNcProgramTransferMain
    {
        public int NcProgramTransferId { get; set; }
        public Nullable<int> McId { get; set; }
        public string ProgramNumber { get; set; }
        public Nullable<int> VersionNumber { get; set; }
        public string ProgramData { get; set; }
        public Nullable<System.DateTime> CreatedDate { get; set; }
        public Nullable<int> CreatedBy { get; set; }
        public Nullable<bool> IsDeleted { get; set; }
        public Nullable<bool> FromCNC { get; set; }
    
        public virtual tblmachinedetail tblmachinedetail { get; set; }
    }
}
