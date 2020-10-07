using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSplivehmiscreen
    {
        public int Sphmiid { get; set; }
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
        public string OperatorDet { get; set; }
        public DateTime? PestartTime { get; set; }
        public int ProcessQty { get; set; }
        public string DdlwokrCentre { get; set; }
        public int IsMultiWo { get; set; }
        public int IsHold { get; set; }
        public string SplitWo { get; set; }
        public int? BatchNo { get; set; }
        public int IsSync { get; set; }
        public string AutoBatchNumber { get; set; }
        public int? ProcessId { get; set; }
        public int? ActivityId { get; set; }
        public int IsBatchStart { get; set; }
        public int? IsActivityFinish { get; set; }
        public int? IsPartialFinish { get; set; }
        public int? ReasonId { get; set; }
        public int BatchHmiid { get; set; }
        public DateTime? ActivityEndTime { get; set; }
        public int? IsGenericClicked { get; set; }
        public int? IsIdleClicked { get; set; }
        public int? IsBreakdownClicked { get; set; }
        public int? IspmClicked { get; set; }
        public int? IsHoldClicked { get; set; }
        public int? IsReworkClicked { get; set; }
        public int IsBatchFinish { get; set; }
        public int? IsChecked { get; set; }
        public string PcpNo { get; set; }
    }
}
