using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.BatchProcessingEntity;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CommonResponseWithMachineName;

namespace DAS.Controllers
{
    [ApiController]
    public class BatchProcessController : Controller
    {
        IBatchProcess batobj;

        #region Monika code saatrts


        public BatchProcessController(IBatchProcess batch)
        {
            batobj = batch;
        }

        //Get DDL List from tblddl where IsCompleted=0
        [Route("BatchProcess/GetDDLList")]
        [HttpPost]
        public async Task<ActionResult> GetDDL(DDLListForBatch data)
        {
            DDLCommonResponseForBatch obj = new DDLCommonResponseForBatch();
            obj = batobj.GetDDLLists(data);
            return Ok(obj);
        }

        //Get WorkCenter Details commented
        [Route("BatchProcess/GetShopWiseWorkCenter")]
        [HttpGet]
        public async Task<ActionResult> GetWC(int MachineID)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.GetShopWiseWorkCenter(MachineID);
            return Ok(obj);
        }

        //Method to Operator ID Auto Suggest 
        [Route("BatchProcess/OperatorIdAutoSuggest")]
        [HttpGet]
        public async Task<ActionResult> GetOpId(string OpID)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.OperatorDetails(OpID);
            return Ok(obj);
        }

        [HttpGet]
        [Route("BatchProcessing/BatchUniqueCodeGenerator")]
        public async Task<IActionResult> BatchUniqueCodeGeneratordetails(string projectName)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = batobj.BatchUniqueCodeGenerator(projectName);
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/NewBatchNumberGenerator")]
        public async Task<IActionResult> NewBatchdetails(int MachineId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = new CommonResponse();
            response = batobj.NewBatchUniqueCodeGenerator(MachineId);
            //return Ok(response);;
            return Ok(response);
        }

        /// <summary>
        ///  Get Machine DisplayName
        /// </summary>
        /// <param name="MachineID"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("BatchProcessing/GetStartedBatch")]
        public async Task<IActionResult> GetStartedBatchdetails(int MachineID)
        {
            //calling DepartmentDAL busines layer
            CommonResponseWithMachinedesscName response = batobj.GetStartedBatch(MachineID);
            return Ok(response);
        }


        //Click of BatchNo Get WO List from tblbactchhmiScreen
        [Route("BatchProcess/GetWoDet")]
        [HttpPost]
        public async Task<ActionResult> GetWoDetails(BatchDetWithMachineName data)
        {
            CommonResponseWithMachinedesscName obj = new CommonResponseWithMachinedesscName();
            obj = batobj.GetWo(data);
            return Ok(obj);
        }

        //get Activity List from tblActivity table
        [Route("BatchProcess/GetActivityList")]
        [HttpGet]
        public async Task<ActionResult> GetActivityListdetails(int machineID)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.GetActivity(machineID);
            return Ok(obj);
        }

        //get Process List from tblProcess table commented
        [Route("BatchProcess/GetProcess")]
        [HttpGet]
        public async Task<ActionResult> GetProcessListdetails()
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.GetProcess();
            return Ok(obj);
        }


        //Click of Start Batch insert into livehmiscreen  commented
        [Route("BatchProcess/StartBatch")]
        [HttpPost]
        public async Task<ActionResult> Getstart(batchdet data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.StartBatch(data);
            return Ok(obj);
        }


        //Set Shift quantity
        [Route("BatchProcess/SetShiftQty")]
        [HttpPost]
        public async Task<ActionResult> SetShiftdetails(SetShiftbatch data)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.SetShift(data);
            return Ok(obj);
        }

        //Set Shift quantity
        [Route("BatchProcess/ChangeShift")]
        [HttpPost]
        public async Task<ActionResult> ChangeShift(string batchNo)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.SetChange(batchNo);
            return Ok(obj);
        }

        //To Get all the batch Number from SPLiveHMISCreen
        [Route("BatchProcess/GetAllWo")]
        [HttpGet]
        public async Task<ActionResult> GetAllWodetails(int MachineID)
        {
            CommonResponseWithMachinedesscName obj = new CommonResponseWithMachinedesscName();
            obj = batobj.GetBatchNo(MachineID);
            return Ok(obj);
        }

        //To Get Select WO of that particular Batch Number  
        [Route("BatchProcess/GetStartedWo")]
        [HttpPost]
        public async Task<ActionResult> GetstartedWodegt(BatchDetWithMachineName data)
        {
            CommonResponseWithMachineName obj = new CommonResponseWithMachineName();
            obj = batobj.GetStartedWo(data);
            return Ok(obj);
        }

        //Set Delivered quantity   
        [Route("BatchProcess/SetDeliveredQty")]
        [HttpPost]
        public async Task<ActionResult> GetDeliveredQtydetails(SetDel data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.SetDeliveredQty(data);
            return Ok(obj);
        }

        //Individual Partail Finish    
        [Route("BatchProcess/IndividualPartialFinish")]
        [HttpPost]
        public async Task<ActionResult> IndividualPartialFinishdetails(ActivityFinish AFData)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.IndividualPartialFinish(AFData);
            return Ok(obj);
        }

        //Individual Activity Finish  
        [Route("BatchProcess/IndividualActivityFinish")]
        [HttpPost]
        public async Task<ActionResult> IndividualActivityFinish(ActivityFinish AFData)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.IndividualActivityFinish(AFData);
            return Ok(obj);
        }

        //Split Wo
        [Route("BatchProcess/SplitWo")]
        [HttpPost]
        public async Task<ActionResult> GetSplitWo(splitWo data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.split(data);
            return Ok(obj);
        }

        //Individual Batch Finish  
        [Route("BatchProcess/IndividualBatchFinish")]
        [HttpPost]
        public async Task<ActionResult> IndividualBatchFinishdetails(ActivityFinish AFData)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.IndividualBatchFinish(AFData);
            return Ok(obj);
        }

        //Click on ReworkOrder 
        [Route("BatchProcess/ReworkOrder")]
        [HttpPost]
        public async Task<ActionResult> Reworkorderdetails(string hmiids)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.SelectReWorkorder(hmiids);
            return Ok(obj);
        }

        //Click on Planned Maintainance Insert into livemodedb
        [Route("BatchProcess/PlannedMaintainance")]
        [HttpGet]
        public async Task<ActionResult> PlannedMaintainancedetails(int MachineId)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.GetPlannedMaintaince(MachineId);
            return Ok(obj);
        }

        //Click on Planned Maintainance Insert into livemodedb   
        [Route("BatchProcess/PlannedMaintain")]
        [HttpPost]
        public async Task<ActionResult> PlannedMaintaindetails(PM data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.Plannedmaintain(data);
            return Ok(obj);
        }

       //on Click of Breakdown Record will insert into tblsplivemode table 
        [Route("BatchProcess/BreakdownEntry")]
        [HttpPost]
        public async Task<ActionResult> BreakdownEntrydetails(PM data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = batobj.BreaksdownEntry(data);
            return Ok(obj);
        }

        //On Click of Generic WO Record will insert into tblSp generic table
        [Route("BatchProcess/GenericWo")]
        [HttpPost]
        public async Task<ActionResult> GenericWodetails(GenericPM data)
        {
            CommonResponse obj = new CommonResponse();
            obj = batobj.GenericWO(data);
            return Ok(obj);
        }

        #endregion monika code ends


        #region Anjali Code satrt

        ////Anjali Code

        ////commented
        [HttpPost]
        [Route("BatchProcessing/UploadDDL")]
        public async Task<IActionResult> UploadDDLdetails(DdlDetails data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = new CommonResponse();
            response = batobj.UploadDDL(data);
            //return Ok(response);;
            return Ok(response);
        }

   

        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("BatchProcessing/AddUploadedDdlBatchDetails")]
        public async Task<IActionResult> AddUploadedDdlBatchDetails([FromBody]BatchDetails data)
        {
            //calling DepartmentDAL busines layer

            AddWOCommonResponse response = new AddWOCommonResponse();

            response = batobj.AddUploadedDdlBatchDetails(data);
            //return Ok(response);;
            return Ok(response);
        }

        [HttpGet]
        [Route("BatchProcessing/GetNewlyUploadedDdlDetails")]
        public async Task<IActionResult> GetNewlyUploadedDdlDetailsdetails(int MachineId)
        {
            //calling DepartmentDAL busines layer
            CommonResponseCountList response = batobj.GetNewlyUploadedDdlDetails(MachineId);
            return Ok(response);
        }

     

        [HttpGet]
        [Route("BatchProcessing/GetModeDetails")]
        public async Task<IActionResult> GetModeDetailsdetails(int MachineId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = batobj.GetModeDetails(MachineId);
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/GetGenericWorkCodes")]
        public async Task<IActionResult> GetGenericWorkCodesdetails(GenericWODetails data)
        {
            //calling DepartmentDAL busines layer
            BreakDownCodeResponse response = new BreakDownCodeResponse();
            response = batobj.GetGenericWorkCodes(data);
            //return Ok(response);;
            return Ok(response);
        }

       

        //[HttpPost]
        //[Route("BatchProcessing/AddBatchDetails")]
        //public async Task<IActionResult> AddBatchDetails([FromBody]BatchDetails data)
        //{
        //    //calling DepartmentDAL busines layer

        //    CommonResponse response = new CommonResponse();

        //    response = batchProcessing.AddBatchDetails(data);
        //    //return Ok(response);;
        //    return Ok(response);
        ////}

        //[HttpGet]
        //[Route("BatchProcessing/GetHoldCodes")]
        //public async Task<IActionResult> GetHoldCodes(int HoldCodeID)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse response = batobj.GetHoldCodes(HoldCodeID);
        //    return Ok(response);
        //}

        [HttpPost]
        [Route("BatchProcessing/GetHoldCodes")]
        public async Task<IActionResult> GetHoldCodesdetails(HoldCodeDetails data)
        {
            //calling DepartmentDAL busines layer
            HoldIdleCodeResponse response = batobj.GetHoldCodes(data);
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/HoldCodeEntry")]
        public async Task<IActionResult> HoldCodeEntrydetails(AddHoldCodes data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse response = new GeneralResponse();
            response = batobj.HoldCodeEntry(data);
            //return Ok(response);;
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/GetBreakDownCodes")]
        public async Task<IActionResult> GetBreakDownCodesdetails(AddBreakDownCodes data)
        {
            //calling DepartmentDAL busines layer
            BreakDownCodeResponse response = batobj.GetBreakDownCodes(data);
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/EndHold")]
        public async Task<IActionResult> EndHolddetails(string hmiiid)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse response = new GeneralResponse();
            response = batobj.EndHold(hmiiid);
            //return Ok(response);;
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/GetIdleCodes")]
        public async Task<IActionResult> GetIdleCodesdetails(BatchProcessingEntity.IdleCodeDetails data)
        {
            //calling DepartmentDAL busines layer
            HoldIdleCodeResponse response = batobj.GetIdleCodes(data);
            return Ok(response);
        }

        [HttpPost]
        [Route("BatchProcessing/IdleLossCodes")]
        public async Task<IActionResult> IdleLossCodesdetails(AddIdleCodes data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = new CommonResponse();
            response = batobj.IdleLossCodes(data);
            //return Ok(response);;
            return Ok(response);
        }

        /// <summary>
        /// Idle List
        /// </summary>
        /// <param name="MachineId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("BatchProcessing/IdleList")]
        public async Task<IActionResult> IdleListdetails(int MachineId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = batobj.IdleList(MachineId);
            return Ok(response);
        }

        /// <summary>
        /// Break Down List
        /// </summary>
        /// <param name="MachineId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("BatchProcessing/BreakDownList")]
        public async Task<IActionResult> BreakDownLists(int MachineId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = batobj.BreakDownList(MachineId);
            return Ok(response);
        }

        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("BatchProcessing/AddBatchDetails")]
        public async Task<IActionResult> AddBatchDetailss([FromBody]BatchDetails data)
        {
            //calling DepartmentDAL busines layer

            AddWOCommonResponse response = new AddWOCommonResponse();

            response = batobj.AddBatchDetails(data);
            //return Ok(response);;
            return Ok(response);
        }

        /// <summary>
        /// Add BatchNo List To Idle
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("BatchProcessing/AddBatchNoListToIdle")]
        public async Task<IActionResult> AddBatchNoListToIdles(AddIdleCodesList data)
        {
            //calling DepartmentDAL busines layer
            CommonResponsewithEror response = new CommonResponsewithEror();
            response = batobj.AddBatchNoListToIdle(data);
            //return Ok(response);;
            return Ok(response);
        }

        /// <summary>
        /// Add BatchNo List To HoldCode
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("BatchProcessing/AddBatchNoListToHoldCode")]
        public async Task<IActionResult> AddBatchNoListToHoldCodes(AddHoldCodesList data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse response = new GeneralResponse();
            response = batobj.AddBatchNoListToHoldCode(data);
            //return Ok(response);;
            return Ok(response);
        }

        #endregion Anjali code end

        [HttpPost]
        [Route("BatchProcessing/OperatorLoginCheck")]
        public async Task<IActionResult> OperatorLoginCheck(Oplogin data)
        {
            //calling DepartmentDAL busines layer
            CommonResponsewithEror response = new CommonResponsewithEror();
            response = batobj.OperatorLoginCheck(data);
            //return Ok(response);;
            return Ok(response);
        }

    }
}
