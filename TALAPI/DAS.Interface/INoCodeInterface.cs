using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
   public interface INoCodeInterface
    {
        //Index get Data
        CommonResponse Index();
        //List<int> GetMachineDet(int CellID);
        EntityModelWithLossCount GetLossDet(lossmodel loss);
        EntityModel GetLossCodeLevel1();
        EntityModel GetLossCodeLevel2(int LossID);
        EntityModel GetLossCodeLevel3(int LossID);
        EntityModel UpdatetcfLoss(updateLoss tcfdata);
        // EntityModel GettcfLoss();
        EntityModel SendToApproveAllLossDetails(EntityHMIDetails data);
        EntityModel UpdateliveLoss(string correctedDate,int MachineID);

        //get Login Details
        CommonResponse1 LoginDetails(LoginInfo data);

        //EntityModel InsertMultiLoss(tcfclass tcfdata, int noofreason);
        //EntityModel GettcfLossLevel(int LossID);
        //EntityModel GetRejecttcfLoss(int id);
        EntityModel GetReason();
        // accepting the sent NoCode
        CommonResponse AcceptAllNoCodeDetails(EntityHMIDetails data);
        // rejecting the sent NoCode
        CommonResponse RejectAllNoCodeDetails(RejectReasonStore data);
        CommonResponse1 NoCodeSplitDurationDetails(List<NoCodeStartEndDateTime> data);
        //Accept Reject NoCode table disaply
        CommonResponse AcceptRejectNoCodeTable(EntityHMIDetails data);

        //if the NoCode details already there means delete the previous record and again insert
        CommonResponse DeletePreviousNoCode(lossmodel loss);
        CommonResponse UpdateTime(CompareDuration data);

        #region split duration

        //Split duration
        CommonResponse SplitDuration(int uawoid);
        //Compare The Duration With End Time
        CommonResponse ComapreEndDuration(CompareDuration data);
        //Confirm the split duration
        CommonResponse ConfirmSpliDuration(ConfirmSplitDuration data);
        //delete the split duration
        CommonResponse DeleteSplitDuration(int uaWOId, string uaWOIds);
        //Clear Records
        CommonResponse ClearAllData(string uaWOIds);

        #endregion
    }
}
