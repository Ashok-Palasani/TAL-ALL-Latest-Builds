using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
    public interface INoLogin
    {
        //Index Data
        CommonResponse Index();
        //Get start and end time from for no login
        CommonResponse1 GetNoLoginTiming(EntityNoLogin data);
        CommonResponse Edit(int nologinid);
        //Login details for Approval
        CommonResponse1 LoginDetails(LoginInfo data);
        //set operatorid
        CommonResponse SetOperatorId(NoLoginOperatorDetails data);
        GeneralResponse1 DeleteSplitDuration(DeleteSplitDuration data);
        CommonResponse1 AfterSplitDuration(string date, int machineId);
        //CommonResponse1 DeleteTempTableData(int noLoginId, int machineId);
        CommonResponse1 DeleteTempTableData(string noLoginIds);
        GeneralResponse1 UpdateSplitDuration(UpdateNoLogin data);
        CommonResponse DeletePreviousNoCode(EntityNoLogin data);
        //set shift
        CommonResponse SetShift(NoLoginSetShift data);
        // getting all the ddl list from tblddl        
        NoLoginDDLCommonResponse GetDDLList(NoLoginDDLList data);
        //Validate DDLList
        CommonResponse ValidateDDLIDS(DDLIds data);
        // listing the shop wise workcenters
        CommonResponse GetShopWiseWorkCenter(int noLoginId);
        //setting rewordorder
        CommonResponse SetReWorkOrder(NoLoginSetReWork data);
        //Check DDL already selectedor not
        bool CheckPrvDDL(int ddlIds);
        // sendging the selected workorder to main screen
        CommonResponse SendWorkOrders(List<NoLoginSelectWO> ddlIds);
        //Split WorkOrder
        CommonResponse SplitWorkOrder(NoLoginSetSplitWork data);
        //Remove the work order
        CommonResponse RemoveWorkOrder(int noLoginId);
        // partial finish the workorder
        CommonResponse PartialFinishWorkOrderDetails(NoLoginStoreHMIDetails data);
        //Job finish the workorder
        CommonResponse JobFinishWorkOrderDetails(NoLoginJobFinish data);
        //set blank the timing
        CommonResponse SetBlank(int nologinId, int value);
        // sending for apporval
        CommonResponse SendToApproveAllWODetails(EntityNoLogin data);
        // accepting the sent workorder
        CommonResponse AcceptAllWODetails(EntityNoLogin data);
        //get reject resons
        CommonResponse GetRejectReason();
        // rejecting the sent workorder
        CommonResponse RejectAllWODetails(RejectReasonStore data);
        //Accept Reject Unassigned WO table disaply
        CommonResponse AcceptRejectUnWOTable(EntityHMIDetails data);

        CommonResponse1 SplitDurationDetails(List<NoLoginStartEndDateTime> data);
        CommonResponse1 GetTimeDetails(int id);

        //CommonResponse1 DeleteTempTableData(int noLoginId, int machineId);
        GeneralResponse1 SplitDuration(NoSplitDetails data);
    }
}
