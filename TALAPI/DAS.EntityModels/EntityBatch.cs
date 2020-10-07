using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class OperatorId
    {
        public string OpId { get; set; }
    }

    public class WoDet
    {
        public int BHmiid { get; set; }
        public string prevBatchNumber { get; set; }
        public string batchNumber { get; set; }
        public string ActivityName { get; set; }
        public string ProcessName { get; set; }
        public string opId { get; set; }
        public string woNo { get; set; }
        public string operationNo { get; set; }
        public string partNo { get; set; }
        public string project { get; set; }
        public string ProdFai { get; set; }
        public string targetQty { get; set; }
        public string ProcessQty { get; set; }
        public string DeliverdQty { get; set; }
        public string pcpNo { get; set; }
    }

    public class GenericWODetails
    {
        public int GenericWorkId { get; set; }
        public int MachineID { get; set; }
        public int Level { get; set; }
    }

    public class getBatchNo
        {
        public string BatchNo { get; set; }
        public bool IsIndividualIdle { get; set; }
        public bool IsIndividualBreakdown { get; set; }
        public bool IsIndividualPM { get; set; }
        public bool IsIndividualHold { get; set; }
    }

    public class GetCreatedBatches
    {
        public string autoBatchNumber { get; set; }
    }

    public class CommonResponseCountList
    {
        public bool isStatus { get; set; }
        public dynamic response { get; set; }
        public string MachineDisplayName { get; set; }
        public int Count { get; set; }
    }

    public class WoDetails
    {
        public int HMIId { get; set; }
        public string ActivityName { get; set; }
        public string ProcessName { get; set; }
        public string woNo { get; set; }
        public string operationNo { get; set; }
        public string partNo { get; set; }
        public string project { get; set; }
        public string opId { get; set; }
        public string ProdFai { get; set; }
        public string targetQty { get; set; }
        public string ProcessQty { get; set; }
        public string DeliverdQty { get; set; }
        public string SplitWo { get; set; }
        public string BatchNo { get; set; }
        public string pcpNo { get; set; }
    }

    public class splitWo
    {
        public int HMIID { get; set; }
        public bool isChecked { get; set; }
    }

    public class DdlData
    {
        public string WorkCenter { get; set; }
        public string ProductionOrder { get; set; }
        public string OperationNo { get; set; }
        public string Operationshortdesc { get; set; }
        public string MatrialDescription { get; set; }
        public string PartNumber { get; set; }
        public string OrderCatgory { get; set; }
        public string TargetQty { get; set; }
        public string MADDate { get; set; }
        public string MADDateInd { get; set; }
        public string PreOpnEndDate { get; set; }
        public string DaysAgeing { get; set; }
        public string Project { get; set; }
        public string DueDate { get; set; }
        public string FiagRushInd { get; set; }
        public string OperationsOnHold { get; set; }
        public string ReasonForHold { get; set; }
        public string SpiltWorkOrder { get; set; }
    }

    public class DdlDetails
    {
        public List<DdlData> ddls { get; set; }
    }

    //public class batchdet
    //{
    //    public int BHmiid { get; set; }
    //    public string WoNo { get; set; }
    //    public string NewAutoBatchNumber { get; set; }
    //    public string OperationNo { get; set; }
    //    public string PartNo { get; set; }
    //    public string project { get; set; }
    //    public string ProdFai { get; set; }
    //    public string OperatorId { get; set; }
    //    public int processId { get; set; }
    //    public string ActivityId { get; set; }
    //    public string TargetQty { get; set; }
    //    public int MachineID { get; set; }
    //    public int ProcessQty { get; set; }
    //    public string SplitWo { get; set; }
    //    public string autoBatchNumber { get; set; }
    //    public int IsChecked { get; set; }
    //}

    public class batchdet
    {
        public string BHmiid { get; set; }
        public string unCheckedId { get; set; }
        public string NewAutoBatchNumber { get; set; }
        public string autoBatchNumber { get; set; }
        public int processId { get; set; }
        public string ActivityId { get; set; }
        public string OperatorId { get; set; }
        public int MachineID { get; set; }
    }



    public class ActivityFinish
    {
        public string HMIId { get; set; }
        public string unCheckedId { get; set; }
        public string autoBatchNumber { get; set; }
        public string activityName { get; set; }
    }
    
    public class breakdoen
    {
        public int Losscodeid { get; set; }
        public string Losscode { get; set; }
    }

    public class Role
    {
        public int Roleid { get; set; }
        public string RoleDesc { get; set; }
    }

    public class process
    {
        public int processid { get; set; }
        public string processdesc { get; set; }
    }

    public class AddIdleCodesList
    {
        public int lossCodeId { get; set; }
        public string batchNo { get; set; }
        public int machineId { get; set; }
        public string status { get; set; }
        public bool IndividualClicked { get; set; }
        public bool OverallClicked { get; set; }
        public int batchNoCount { get;set;}
    }

    public class AddHoldCodesList
    {
        public string BatchNo { get; set; }
        public int HoldCodeId { get; set; }
        public int MachineId { get; set; }
        public string Color { get; set; }
    }

    public class IdleCodeDetails
    {
        public int LossCodeID { get; set; }
        public int MachineID { get; set; }
        public int Level { get; set; }
        public bool IsTrue { get; set; }
    }

    public class Activity
    {
        public int ActivityId { get; set; }
        public string ActivityName { get; set; }
    }

    public class SetDel
    {
        public int HMIID { get; set; }
        public int DeliveredQty { get; set; }
    }

    public class SetShiftbatch
    {
        public string BatchNo { get; set; }
        public string Shift { get; set; }
    }

    public class PM
    {
        public int MachineID { get; set; }
        public int LossCodeId { get; set; }
        public string Status { get; set; }
        public string BatchNo { get; set; }
        public bool IndividualClicked { get; set; }
        public bool OverallClicked { get; set; }
    }

    public class GenericPM
    {
        public int MachineID { get; set; }
        public int LossCodeId { get; set; }
        public string Status { get; set; }
        public string BatchNo { get; set; }
        public bool IndividualClicked { get; set; }
        public bool OverallClicked { get; set; }
        public string opId { get; set; }
    }

    public class HoldCodeDetails
    {
        public int HoldCodeID { get; set; }
        public int MachineID { get; set; }
        public int Level { get; set; }
        public bool IsTrue { get; set; }
    }

    //// For DDl Lasy Loading
    //public class DDLList
    //{
    //    public int machineId { get; set; }
    //    public int takeValue { get; set; }
    //    public int skipeValue { get; set; }
    //    //public int count { get; set; }
    //}
}
