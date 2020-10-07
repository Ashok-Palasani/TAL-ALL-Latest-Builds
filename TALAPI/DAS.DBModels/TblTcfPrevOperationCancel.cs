using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblTcfPrevOperationCancel
    {
        public int TcfopcancelId { get; set; }
        public string ProductionOrder { get; set; }
        public string Operation { get; set; }
        public int IsCancelled { get; set; }
        public int? Qty { get; set; }
        public string CorrectedDate { get; set; }
        public string WorkCenter { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public int SendApprove { get; set; }
        public int AcceptReject { get; set; }
        public int ApprovalLevel { get; set; }
        public int Update { get; set; }
        public string PartNumber { get; set; }
        public string UploadDate { get; set; }
        public int AcceptReject1 { get; set; }
        public int RejectReason { get; set; }
        public int RejectReason1 { get; set; }
        public int UpdateLevel { get; set; }
        public int IsPending { get; set; }
    }
}
