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
    
    public partial class tbllivemanuallossofentry
    {
        public int MLossID { get; set; }
        public int MessageCodeID { get; set; }
        public Nullable<System.DateTime> StartDateTime { get; set; }
        public string CorrectedDate { get; set; }
        public int MachineID { get; set; }
        public string Shift { get; set; }
        public string MessageDesc { get; set; }
        public string MessageCode { get; set; }
        public Nullable<int> HMIID { get; set; }
        public string PartNo { get; set; }
        public Nullable<int> OpNo { get; set; }
        public string WONo { get; set; }
        public Nullable<System.DateTime> EndDateTime { get; set; }
        public Nullable<int> EndHMIID { get; set; }
        public Nullable<System.DateTime> DeletedDate { get; set; }
    
        public virtual tbllossescode tbllossescode { get; set; }
    }
}
