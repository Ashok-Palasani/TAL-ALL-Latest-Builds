using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.BatchProcessingEntity;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CommonResponseWithMachineName;

namespace DAS.Interface
{
   public interface IBatchProcess
    {
        DDLCommonResponseForBatch GetDDLLists(DDLListForBatch data);
        CommonResponse GetShopWiseWorkCenter(int machineId);
        CommonResponseWithMachinedesscName GetWo(BatchDetWithMachineName data);
        CommonResponsewithEror StartBatch(batchdet data);
        CommonResponseWithMachineName GetStartedWo(BatchDetWithMachineName data);
        CommonResponseWithMachinedesscName GetBatchNo(int MachineID);
        CommonResponsewithEror IndividualActivityFinish(ActivityFinish AFData);
        CommonResponsewithEror IndividualBatchFinish(ActivityFinish AFData);
        CommonResponsewithEror IndividualPartialFinish(ActivityFinish AFData);
        CommonResponse SelectReWorkorder(string hmiids);
        CommonResponsewithEror GetPlannedMaintaince(int MachineID);
        CommonResponsewithEror Plannedmaintain(PM data);
        CommonResponsewithEror BreaksdownEntry(PM data);
        CommonResponse GenericWO(GenericPM data);
        CommonResponse GetActivity(int machineID);
        CommonResponse GetProcess();
        CommonResponsewithEror SetDeliveredQty(SetDel data);
        CommonResponse SetShift(SetShiftbatch data);
        CommonResponse SetChange(string batchNo);
        CommonResponsewithEror split(splitWo data);
        CommonResponse OperatorDetails(string OpID);
        //Anjali COde

        AddWOCommonResponse AddUploadedDdlBatchDetails(BatchDetails data);
        CommonResponse UploadDDL(DdlDetails data);
        CommonResponse NewBatchUniqueCodeGenerator(int MachineId);
        CommonResponseCountList GetNewlyUploadedDdlDetails(int MachineId);
        BreakDownCodeResponse GetGenericWorkCodes(GenericWODetails data);
        CommonResponse BatchUniqueCodeGenerator(string projectName);
        //CommonResponse AddBatchDetails(BatchDetails data);
        //CommonResponse GetHoldCodes(int HoldCodeID);
        GeneralResponse HoldCodeEntry(AddHoldCodes data);
        GeneralResponse EndHold(string hmiiid);
        //CommonResponse GetIdleCodes(int LossCodeID, bool isStart);
        CommonResponse IdleLossCodes(AddIdleCodes data);
        CommonResponse IdleList(int MachineId = 0);
        CommonResponse BreakDownList(int MachineId = 0);
        AddWOCommonResponse AddBatchDetails(BatchDetails data);
        HoldIdleCodeResponse GetHoldCodes(HoldCodeDetails data);
        HoldIdleCodeResponse GetIdleCodes(BatchProcessingEntity.IdleCodeDetails data);
        CommonResponse GetModeDetails(int MachineId);
        CommonResponseWithMachinedesscName GetStartedBatch(int MachineID);
        BreakDownCodeResponse GetBreakDownCodes(AddBreakDownCodes data);
        CommonResponsewithEror AddBatchNoListToIdle(AddIdleCodesList data);
        GeneralResponse AddBatchNoListToHoldCode(AddHoldCodesList data);
    }
}
