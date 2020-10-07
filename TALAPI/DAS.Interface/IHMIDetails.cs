using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
   public interface IHMIDetails
    {
        //Index data
        CommonResponse Index();
        // get the details from HMI screen and insert into tbluasigned
        CommonResponse GetMachineDetailsHMI(EntityHMIDetails data);
        //Login details for Approval
        CommonResponse1 LoginDetails(LoginInfo data);
        //To Edit the Current Data
        CommonResponse Edit(int uawoid);
        //set operatorid
        CommonResponse SetOperatorId(OperatorDetails data);
        CommonResponse DeletePreviousNoCode(EntityHMIDetails data);
        //set shift
        CommonResponse SetShift(SetShift data);
        //Split duration
        CommonResponse SplitDuration(int uawoid);
        //Compare The Duration With End Time
        CommonResponse ComapreEndDuration(CompareDuration data);
        //Confirm the split duration
        CommonResponse ConfirmSpliDuration(ConfirmSplitDuration data);
        //delete the split duration
        CommonResponse DeleteSplitDuration(int uaWOId,string uaWOIds);
        //Clear Records
        CommonResponse ClearAllData(string uaWOIds);
        // getting all the ddl list from tblddl        
        DDLCommonResponse GetDDLList(DDLList data);
        // listing the shop wise workcenters
        CommonResponse GetShopWiseWorkCenter(int uaWoID);
        // sendging the selected workorder to main screen
        CommonResponse SendWorkOrders(List<SelectWO> ddlIds);
        //Check DDL already selectedor not
        bool CheckPrvDDL(int ddlIds);
        //Start The workOrder
        CommonResponse StartWO(int uaWoID);
        //Get Hold Code by Levels        
        CommonResponse GetHoldCodes(HoldCode data);
        //Set The hold Code Reasons
        CommonResponse SetHoldCodereason(HoldCodeGet data);
        //Get Generic Code by Levels
        CommonResponse GetGenericCodes(GenericCode data);
        //Set The Generic Code Reasons
        CommonResponse SetGenericCodereason(GenericCodeGet data);
        //Remove the work order
        CommonResponse RemoveWorkOrder(int uawoid);
        //Validate DDLList
        CommonResponse ValidateDDLIDS(DDLIds data);
        //Split WorkOrder
        CommonResponse SplitWorkOrder(SetSplitWork data);
        // job finish the workorder
        CommonResponse JobFinishWorkOrderDetails(StoreHMIDetails data);
        // partial finish the workorder
        CommonResponse PartialFinishWorkOrderDetails(StoreHMIDetails data);
        //set blank the timing
        CommonResponse SetBlank(int unwoid,int value);
        //setting rewordorder
        CommonResponse SetReWorkOrder(SetReWork data);
        // sending for apporval
        CommonResponse SendToApproveAllWODetails(EntityHMIDetails data);
        // accepting the sent workorder
        CommonResponse AcceptAllWODetails(EntityHMIDetails data);
        //backup Report tables
        bool TakeBackupReportData(string correctedDate);
        //Get Reject Reasons
        CommonResponse GetRejectReason();
        // rejecting the sent workorder
        CommonResponse RejectAllWODetails(RejectReasonStore data);
        //Accept Reject Unassigned WO table disaply
        CommonResponse AcceptRejectUnWOTable(EntityHMIDetails data);
        // updating for report tables
        CommonResponse UpdateToReportTables(string correctedDate,int MachineID);
        // For Generic Workorder
        CommonResponse GenericWorkOrder(GenericWO genericWO);
        CommonResponse1 UnassignedWOSplitDurationDetails(List<UnassignedWOStartEndDateTime> data);

        CommonResponse UpdateTime(CompareDuration data);

    }
}
