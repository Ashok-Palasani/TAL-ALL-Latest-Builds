using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class EntityHMIDetails
    {
        public int plantId { get; set; }
        public int shopiId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public string fromDate { get; set; }
        public string id { get; set; }
        public string unCheckId { get; set; }
        //public string toDate { get; set; }  // for selecting multiple day
    }

    // for sending the DDL list to View
    public class EDDLList
    {
        public int ddlId { get; set; }
        public string project { get; set; }
        public string woNo { get; set; }
        public string partNo { get; set; }
        public string opNo { get; set; }
        public string targetQty { get; set; }
        public string daysAgeing { get; set; }
        public string flagRush { get; set; }
        public string madDateInd { get; set; }
        public string madDate { get; set; }
        public string splitWO { get; set; }
        public int slno { get; set; }
        public string pcpNo { get; set; }
    }

    // to set the operatorid
    public class OperatorDetails
    {
        public int operatorId { get; set; }
        public int uawoid { get; set; }
    }
    // to set the Shift
    public class SetShift
    {
        public string shift { get; set; }
        public int uawoid { get; set; }
    }

    // for storing workcenter
    public class EShopWiseWC
    {
        public int machineId { get; set; }
        public string workCenterName { get; set; }
    }

    //Sending the DDLids and new table primaryKey
    public class SendDDLUnAsignedWoId
    {
        public int unAsignedId { get; set; }
        public string ddlIds { get; set; }
    }

    // for selecting workorder
    public class SelectWO
    {
        public string ddlId { get; set; }
        public int uaWOId { get; set; }
    }

    //Hold code send value
    public class HoldCode
    {
        public int holdCodeId { get; set; }
        public string holdCodeName { get; set; } 
        public int holdCodeLevel { get; set; }
    }

    //Hold Code set Value
    public class HoldCodeGet
    {
        public int holdCodeId { get; set; }
        public int uaWOId { get; set; }
    }

    //Generic code send value
    public class GenericCode
    {
        public int genericCodeId { get; set; }
        public string genericCodeName { get; set; }
        public int genericCodeLevel { get; set; }
    }

    //generic Code set Value
    public class GenericCodeGet
    {
        public int genericCodeId { get; set; }
        public int uaWOId { get; set; }
    }


    // fro storing form ddl to new table
    public class StoreToUnAsigned
    {
        public int uaWOId { get; set; }
        public string workOrderNo { get; set; }
        public string partNo { get; set; }
        public string oprationNo { get; set; }
        public string project { get; set; }
        public string workOrderQty { get; set; }
        public string ddlWorkCenter { get; set; }
        public string prodFai { get; set; }
    }

    // after storing the WorkOrder sending to HMI Screen
    public class SendToHMIScreen
    {
        public int uaWOId { get; set; }
        public string machineName { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string workOrderNo { get; set; }
        public string partNo { get; set; }
        public string oprationNo { get; set; }
        public string workOrderQty { get; set; }
        public string project { get; set; }
    }


    // for generic workorder
    public class GenericWO
    {
        public int uaWOId { get; set; }
    }

    // For DDl Lasy Loading
    public class DDLList
    {
        public int machineId { get; set; }
        public int takeValue { get; set; }
        public int skipeValue { get; set; }
        public int uaWoID { get; set; }
        //public int count { get; set; }
    }

    public class DDLListForBatch
    {
        public int machineId { get; set; }
        public int takeValue { get; set; }
        public int skipeValue { get; set; }
        //public int count { get; set; }
    }

    // Send Reject Reason
    public class RejectReasonData
    {
        public int reasonId { get; set; }
        public string reasonName { get; set; }
    }

    // store reject reason for particular date
    public class RejectReasonStore
    {
        public string correctedDate { get; set; }
        public int reassonId { get; set; }
        public int plantId { get; set; }
        public int shopiId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public string id { get; set; }
        public string unCheckId { get; set; }
    }

    // Shop wise Work Center
    public class ShopWiseWC
    {
        public int uaWoId { get; set; }
        //public int machineId { get; set; }
    }

    public class DDLCommonResponse
    {
        public bool isTure { get; set; }
        public dynamic response { get; set; }
        public int count { get; set; }
    }

    public class DDLCommonResponseForBatch
    {
        public bool isTure { get; set; }
        public dynamic response { get; set; }
        public int count { get; set; }
        public string MachineDispName { get; set; }
        public string MachineInvNo { get; set; }
    }

    public class AddModes
    {
        public int Uawoid { get; set; }
        public int machineId { get; set; }
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
    }

    public class UpdateModes
    {
        public int Uawoid { get; set; }
        public int modeId { get; set; }
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
    }

    public class MachineHmiDet
    {
        public int machineId { get; set; }
        public string machineName { get; set; }
        public List<SendHMIDetails> hmiDetails { get; set; }
    }

    public class SendHMIDetails
    {
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string correctedDate { get; set; }
        public int machineId { get; set; }
    }

    public class StoreHMIDetails
    {
        public int uaWOId { get; set; }
        public string deliveryQty { get; set; }
    }

    //For validating the DDLIDS
    public class DDLIds
    {
        public string ddlIds { get; set; }
    }

    // for set reworkorder
    public class SetReWork
    {
        public int uaWOId { get; set; }
        public int isChecked { get; set; }
    }

    // for set Splitworkorder
    public class SetSplitWork
    {
        public int uaWOId { get; set; }
        public int isChecked { get; set; }
    }

    // for set SplitDuartion
    //public class SplitDuration
    //{
    //    public string startTime { get; set; }
    //    public string endDate { get; set; }
    //    public string endHour { get; set; }
    //    public string endMinute { get; set; }
    //    public string endSecond { get; set; }
    //    public int uaWOId { get; set; }
    //}



    // for set CompareDuartion
    //public class CompareDuration
    //{
    //    public string endDate { get; set; }
    //    public string endHour { get; set; }
    //    public string endMinute { get; set; }
    //    public string endSecond { get; set; }
    //    public int uaWOId { get; set; }
    //}

    // for set SplitDuartion
    public class SplitDuration
    {
        public string startTime { get; set; }
        public string endTime { get; set; }
        //public string endHour { get; set; }
        //public string endMinute { get; set; }
        //public string endSecond { get; set; }
        public int uaWOId { get; set; }
    }

    // for set SplitDuartion
    public class SplitDurationList
    {
        public string uaWOIds { get; set; }
        public List<SplitDuration> listSplitDuration { get; set; }
    }


    // for set CompareDuartion and confiramation
    public class CompareDuration
    {
        public string endTime { get; set; }
        //public string endHour { get; set; }
        //public string endMinute { get; set; }
        //public string endSecond { get; set; }
        public int uaWOId { get; set; }
        public string uaWOIdS { get; set; }
    }

    //Confirm split duration
    public class ConfirmSplitDuration
    {
        public string uaWOIds { get; set; }
        public int flageAcceptReject { get; set; }
    }


    public class UnsignedWO
    {
        public int uaWOId { get; set; }
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string machineName { get; set; }
        public string workOrderNo { get; set; }
        public string partNo { get; set; }
        public string oprationNo { get; set; }
        public string workOrderQty { get; set; }
        public string project { get; set; }
        public string processedQty { get; set; }
        public string prodFai { get; set; }
        public string operatorId { get; set; }
        public string deleveredQty { get; set; }
        public string shift { get; set; }
        public int isWocenter { get; set; }
        public bool ddl { get; set; }
        public bool jf { get; set; }
        public bool pf { get; set; }
        public bool rwo { get; set; }
        public bool isSplit { get; set; }
        //public bool start { get; set; }
    }

    //Index get all data
    public class UnsignedWOData
    {
        public int uaWOId { get; set; }
        //public string startDateTime { get; set; }
        //public string endDateTime { get; set; }
        public string machineName { get; set; }
        //public string workOrderNo { get; set; }
        //public string partNo { get; set; }
        //public string oprationNo { get; set; }
        //public string workOrderQty { get; set; }
        //public string project { get; set; }
        //public string processedQty { get; set; }
        //public string prodFai { get; set; }
        //public string operatorId { get; set; }
        //public string deleveredQty { get; set; }
        //public string shift { get; set; }
        //public int isWocenter { get; set; }
        //public bool ddl { get; set; }
        //public bool pf { get; set; }
        //public bool rwo { get; set; }
        //public bool isSplit { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string CorrectedDate { get; set; }
        public string firstApproval { get; set; }
        public string secondApproval { get; set; }
        public int machineId { get; set; }
    }

    public class UnassignedWOStartEndDateTime
    {
        public int UnassignedWoId { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string message { get; set; }
    }

    public class HMIDetails
    {
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string correctedDate { get; set; }
        public int machineId { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string machineName { get; set; }
        public double duration { get; set; }
    }
}
