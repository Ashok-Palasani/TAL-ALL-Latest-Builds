using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblunasignedwo
    {
        public int Uawoid { get; set; }
        public DateTime Starttime { get; set; }
        public DateTime Endtime { get; set; }
        public int Isdeleted { get; set; }
        public string Correcteddate { get; set; }
        public string Createdon { get; set; }
        public int Createdby { get; set; }
        public string WorkOrderNo { get; set; }
        public string PartNo { get; set; }
        public string OprationNo { get; set; }
        public string Project { get; set; }
        public string ProdFai { get; set; }
        public string WorkOrderQty { get; set; }
        public string ProcessedQty { get; set; }
        public string DeliveryQty { get; set; }
        public int SendApprove { get; set; }
        public int Acceptreject { get; set; }
        public int Isupdate { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string Shiftid { get; set; }
        public string Operatorid { get; set; }
        public int? Machineid { get; set; }
        public int? JobFinish { get; set; }
        public int? PartialFinish { get; set; }
        public int? GenericWo { get; set; }
        public int? ReWork { get; set; }
        public int? IsSplit { get; set; }
        public int? Loperatorid { get; set; }
        public int? Status { get; set; }
        public int? Isworkinprogress { get; set; }
        public int? Isworkorder { get; set; }
        public DateTime? Pestarttime { get; set; }
        public string Ddlworkcenter { get; set; }
        public int? Holdcodeid { get; set; }
        public string Holdcodereason { get; set; }
        public int Rejectreasonid { get; set; }
        public int GenericCodeid { get; set; }
        public string GenericCodereason { get; set; }
        public int IsStart { get; set; }
        public int IsSplitDuration { get; set; }
        public int ApprovalLevel { get; set; }
        public int AcceptReject1 { get; set; }
        public int RejectReason { get; set; }
        public int RejectReason1 { get; set; }
        public int UpdateLevel { get; set; }
        public int IsBlank { get; set; }
        public int IsPending { get; set; }
    }
}
