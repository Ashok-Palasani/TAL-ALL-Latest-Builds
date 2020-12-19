//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CalOEEDaily
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblhmiscreen
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public tblhmiscreen()
        {
            this.tbl_multiwoselection = new HashSet<tbl_multiwoselection>();
        }
    
        public int HMIID { get; set; }
        public int MachineID { get; set; }
        public int OperatiorID { get; set; }
        public string Shift { get; set; }
        public Nullable<System.DateTime> Date { get; set; }
        public Nullable<System.DateTime> Time { get; set; }
        public string Project { get; set; }
        public string PartNo { get; set; }
        public string OperationNo { get; set; }
        public Nullable<int> Rej_Qty { get; set; }
        public string Work_Order_No { get; set; }
        public Nullable<int> Target_Qty { get; set; }
        public Nullable<int> Delivered_Qty { get; set; }
        public Nullable<int> Status { get; set; }
        public string CorrectedDate { get; set; }
        public string Prod_FAI { get; set; }
        public int isUpdate { get; set; }
        public int DoneWithRow { get; set; }
        public int isWorkInProgress { get; set; }
        public int isWorkOrder { get; set; }
        public string OperatorDet { get; set; }
        public Nullable<System.DateTime> PEStartTime { get; set; }
        public int ProcessQty { get; set; }
        public string DDLWokrCentre { get; set; }
        public int IsMultiWO { get; set; }
        public int IsHold { get; set; }
        public string SplitWO { get; set; }
        public Nullable<int> HMIMonth { get; set; }
        public Nullable<int> HMIYear { get; set; }
        public Nullable<int> HMIWeekNumber { get; set; }
        public Nullable<int> HMIQuarter { get; set; }
        public Nullable<int> batchCount { get; set; }
        public int IsSync { get; set; }
        public Nullable<int> SheetNoList { get; set; }
        public Nullable<int> isSplitSapUpdated { get; set; }
        public int prevQty { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tbl_multiwoselection> tbl_multiwoselection { get; set; }
        public virtual tblmachinedetail tblmachinedetail { get; set; }
    }
}
