using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CommonResponseWithMachineName;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
    public interface ISplitDuration
    {
        CommonResponse1 GetPlants();
        CommonResponse1 GetShops(int plantId);
        CommonResponse1 GetCells(int shopId);
        CommonResponse1 GetMachines(int cellId);
        CommonResponse1 GetModeDetails(int MachineId, string date);
        CommonResponse1 ValidateEndtime(int ModeId, string EndTime);
        CommonResponse1 AddModeDetailsToTempTable(List<ModeStartEndDateTime> data);
        CommonResponse1 LoginDetails(LoginInfo data);
        CommonResponse1 AddModeDetails(string tempModeIds);
        //CommonResponse UpdateModeDetails(List<ModeStartEndDateTime> data);
        //CommonResponse1 AddLossDetails(List<AddReasonsDetails> data);
        CommonResponse1 AddLossDetails(int tempModeId);
        CommonResponse1 GetIdleResonLevel1(string mode);
        CommonResponse1 GetIdleResonLevel2(int LossCodeID);
        CommonResponse1 GetIdleResonLevel3(int LossCodeID);
        //CommonResponse1 GetBreakDownReasonLevel1();
        //CommonResponse1 GetBreakDownReasonLevel2(int LossCodeID);
        //CommonResponse1 GetBreakDownReasonLevel3(int LossCodeID);
        //CommonResponse1 DeleteSplitDuration(int tempModeId, int modeId);
        //CommonResponse1 DeleteTempTableData(int machineId,int modeId);
        CommonResponse1 ApprovedSendToMainModeDetails(string tempModeIds);
        CommonResponse1 SendMailToApproveOrReject(SendMailDetails data);
        CommonResponse1 GetRejectReasons();
        CommonResponse1 SendMailAfterReject(SendMailDetails data);
        CommonResponse1 AfterSplitDuration(int tempModeId);
        CommonResponse1 GetApprovedDetails(SendMailDetails data);
        CommonResponse1 Index();
        CommonResponseWithTempId AddIndividualModeDetails(AddIndividualMode data);
        CommonResponseWithTempId GetReasonsForThirdLevel(int tempModeId, int lossCodeId);
        CommonResponse1 GetTimeDetailsForSplitDuration(int tempModeId);
        CommonResponse1 Index1();
        GeneralResponse1 SplitDuration(SplitDetails data);
        //CommonResponse1 DeleteTempModeIds(DeleteTempMode data);
        //CommonResponse1 SplitDurationDetails(List<NoLoginStartEndDateTime> data);
        //CommonResponse1 NoCodeSplitDurationDetails(List<NoCodeStartEndDateTime> data);
        //CommonResponse1 UnassignedWOSplitDurationDetails(List<UnassignedWOStartEndDateTime> data);
        //CommonResponse1 GetTimeDetails(int id);
        //CommonResponse1 GetTimeDetailsForNoCode(int id);
        //CommonResponse1 GetTimeDetailsForUnassignedWo(int id);

        CommonResponse1 DeleteTempTableData(string tempmodeIds);

        GeneralResponse1 DeleteTempModeIds(DeleteTempMode data);
        GeneralResponse1 UpdateSplitDuration(UpdateTempMode data);
        CommonResponse1 UpdateLosses(int tempodeId);
        //CommonResponse1 GetApprovedInfo(string tempModeIds);

    }
}
