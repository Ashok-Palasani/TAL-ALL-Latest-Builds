using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblbatchhmiscreen
    {
        public int Bhmiid { get; set; }
        public int MachineId { get; set; }
        public string OperatiorId { get; set; }
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
        public DateTime? PestartTime { get; set; }
        public int ProcessQty { get; set; }
        public string DdlwokrCentre { get; set; }
        public int IsHold { get; set; }
        public string SplitWo { get; set; }
        public string AutoBatchNumber { get; set; }
        public int? ActivityId { get; set; }
        public int IsBatchStart { get; set; }
        public int? IsActivityFinish { get; set; }
        public int? IsBatchFinish { get; set; }
        public int IsChecked { get; set; }
        public int? ProcessId { get; set; }
        public string PrevBatchNo { get; set; }
        public int? IsPatialFinish { get; set; }
        public int? IsGenericClicked { get; set; }
        public string PcpNo { get; set; }
    }
}
