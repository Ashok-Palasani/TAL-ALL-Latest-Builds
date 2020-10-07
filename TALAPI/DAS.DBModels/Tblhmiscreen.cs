using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblhmiscreen
    {
        public Tblhmiscreen()
        {
            TblMultiwoselection = new HashSet<TblMultiwoselection>();
        }

        public int Hmiid { get; set; }
        public int MachineId { get; set; }
        public int OperatiorId { get; set; }
        public string Shift { get; set; }
        public DateTime? Date { get; set; }
        public DateTime? Time { get; set; }
        public string Project { get; set; }
        public string PartNo { get; set; }
        public string OperationNo { get; set; }
        public int? RejQty { get; set; }
        public string WorkOrderNo { get; set; }
        public int? TargetQty { get; set; }
        public int? DeliveredQty { get; set; }
        public int? Status { get; set; }
        public string CorrectedDate { get; set; }
        public string ProdFai { get; set; }
        public int IsUpdate { get; set; }
        public int DoneWithRow { get; set; }
        public int IsWorkInProgress { get; set; }
        public int IsWorkOrder { get; set; }
        public string OperatorDet { get; set; }
        public DateTime? PestartTime { get; set; }
        public int ProcessQty { get; set; }
        public string DdlwokrCentre { get; set; }
        public int IsMultiWo { get; set; }
        public int IsHold { get; set; }
        public string SplitWo { get; set; }
        public int? Hmimonth { get; set; }
        public int? Hmiyear { get; set; }
        public int? HmiweekNumber { get; set; }
        public int? Hmiquarter { get; set; }
        public int? BatchCount { get; set; }
        public int IsSync { get; set; }
        public int? SheetNoList { get; set; }
        public string AutoBatchNumber { get; set; }
        public string ActivityName { get; set; }
        public int IsBatchStart { get; set; }
        public int BatchHmiid { get; set; }
        public int Pcsflag { get; set; }
        public string OnHoldFlag { get; set; }
        public int OnHoldCode { get; set; }
        public int IsSplitSapUpdated { get; set; }

        public virtual Tblmachinedetails Machine { get; set; }
        public virtual ICollection<TblMultiwoselection> TblMultiwoselection { get; set; }
    }
}
