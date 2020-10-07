using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
  public  class OPCancelEntity
    {
        public class OPCancelDetails
        {
            //public int OpcancelId { get; set; }
            public string ProductionOrder { get; set; }
            public string Operation { get; set; }
            public int IsCancelled { get; set; }
            public int? ProcessedQty { get; set; }
            public string CorrectedDate { get; set; }
            public string WorkCenter { get; set; }
        }
        public class LsitOPcancelDet
        {
            public List<OPCancelDetails> OPCancelDetailsList { get; set; }
        }

        //for Approval
        public class ApprovalClass
        {
            public string id { get; set; }
            public string unCheckId { get; set; }
            public string uploadDate { get; set; }
        }

        //for Reject
        public class RejectClass
        {
            public string id { get; set; }
            public string unCheckId { get; set; }
            public int reasonId { get; set; }
            public string uploadDate { get; set; }
        }

        public class LsitOPcancelDetSucc
        {
            public List<OPCancelDetails> OPCancelDetailsSuccessList { get; set; }
            public List<OPCancelDetails> OPCancelDetailsErrorList { get; set; }
        }

        //for cenclation operation number
        public class OperationCancelDet
        {
            public int opCancelID { get; set; }
            public string productionOrder { get; set; }
            public string operation { get; set; }
            public int isCancelled { get; set; }
            public int processedQty { get; set; }
            public string correctedDate { get; set; }
            public string workCenter { get; set; }
            public string partNo { get; set; }
        }

        //for indexing and display
        //for cenclation operation number
        public class IndexOperationCancelDet
        {
            public int opCancelID { get; set; }
            public string productionOrder { get; set; }
            public string operation { get; set; }
            public int isCancelled { get; set; }
            public int processedQty { get; set; }
            public string correctedDate { get; set; }
            public string workCenter { get; set; }
            public string partNo { get; set; }
            public string sentApp { get; set; }
            public string acceptReject { get; set; }
        }

        //for cenclation operation number
        public class IndexOperationCancelDet1
        {
            public int opCancelID { get; set; }
            public string productionOrder { get; set; }
            public string operation { get; set; }
            public int isCancelled { get; set; }
            public int processedQty { get; set; }
            public string correctedDate { get; set; }
            public string workCenter { get; set; }
            public string partNo { get; set; }
            public string firstApproval { get; set; }
            public string secondApproval { get; set; }
        }

        //returning success and error list 
        public class SELOperationCancelDet
        {
            public List<OperationCancelDet> OPCancelDetailsSuccessList { get; set; }
            public List<OperationCancelDet> OPCancelDetailsErrorList { get; set; }
        }

        //Accept the check box value
        public class AcceptCancel
        {
            public int opCancelID { get; set; }
            public int isChecked { get; set; }
        }

        //Getting data from excel file
        public class UpLoadExcel
        {
            public string WorkCenter { get; set; }
            public string ProductionOrder { get; set; }
            public string PartNumber { get; set; }
            public string Operation { get; set; }
            public int Qty { get; set; }
            public string CorrectedDate { get; set; }
        }
    }

}
