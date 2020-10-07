using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CommonResponseWithMachineName;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class SplitDurationController : ControllerBase
    {
        private readonly ISplitDuration splitDuration;
        public SplitDurationController(ISplitDuration _splitDuration)
        {
            splitDuration = _splitDuration;
        }

        /// <summary>
        /// Login Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/LoginDetails")]
        public async Task<IActionResult> LoginDetails(LoginInfo data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.LoginDetails(data);
            return Ok(response);
        }

        /// <summary>
        /// Get Plants
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetPlants")]
        public async Task<IActionResult> GetPlants()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetPlants();
            return Ok(response);
        }

        /// <summary>
        /// Get Shops
        /// </summary>
        /// <param name="plantId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetShops")]
        public async Task<IActionResult> GetShops(int plantId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetShops(plantId);
            return Ok(response);
        }

        /// <summary>
        /// Get Cells
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetCells")]
        public async Task<IActionResult> GetCells(int shopId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetCells(shopId);
            return Ok(response);
        }

        /// <summary>
        /// Get Machines
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetMachines")]
        public async Task<IActionResult> GetMachines(int shopId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetMachines(shopId);
            return Ok(response);
        }

        /// <summary>
        /// Get Mode Details
        /// </summary>
        /// <param name="MachineId"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetModeDetails")]
        public async Task<IActionResult> GetModeDetails(int MachineId, string date)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetModeDetails(MachineId, date);
            return Ok(response);
        }

        /// <summary>
        /// Validate End time
        /// </summary>
        /// <param name="ModeId"></param>
        /// <param name="EndTime"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/ValidateEndtime")]
        public async Task<IActionResult> ValidateEndtime(int ModeId, string EndTime)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.ValidateEndtime(ModeId, EndTime);
            return Ok(response);
        }

        /// <summary>
        /// Add Mode Details To Temp Table before clicking save button
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/AddModeDetailsToTempTable")]
        public async Task<IActionResult> AddModeDetailsToTempTable(List<ModeStartEndDateTime> data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.AddModeDetailsToTempTable(data);
            return Ok(response);
        }

        /// <summary>
        /// After clicking save button Add Mode Details
        /// </summary>
        /// <param name="tempModeIds"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/AddModeDetails")]
        public async Task<IActionResult> AddModeDetails(string tempModeIds)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.AddModeDetails(tempModeIds);
            return Ok(response);
        }

        //[HttpPost]
        //[Route("SplitDuration/UpdateModeDetails")]
        //public async Task<IActionResult> UpdateModeDetails(List<ModeStartEndDateTime> data)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse response = splitDuration.UpdateModeDetails(data);
        //    return Ok(response);
        //}

        /// <summary>
        /// After adding loss reasons save
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/AddLossDetails")]
        public async Task<IActionResult> AddLossDetails(int tempModeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.AddLossDetails(tempModeId);
            return Ok(response);
        }

        [HttpGet]
        [Route("SplitDuration/GetIdleResonLevel1")]
        public async Task<IActionResult> GetIdleResonLevel1(string mode)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetIdleResonLevel1(mode);
            return Ok(response);
        }

        [HttpGet]
        [Route("SplitDuration/GetIdleResonLevel2")]
        public async Task<IActionResult> GetIdleResonLevel2(int LossCodeID)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetIdleResonLevel2(LossCodeID);
            return Ok(response);
        }

        [HttpGet]
        [Route("SplitDuration/GetIdleResonLevel3")]
        public async Task<IActionResult> GetIdleResonLevel3(int LossCodeID)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetIdleResonLevel3(LossCodeID);
            return Ok(response);
        }

        //[HttpGet]
        //[Route("SplitDuration/GetBreakDownReasonLevel1")]
        //public async Task<IActionResult> GetBreakDownReasonLevel1()
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.GetBreakDownReasonLevel1();
        //    return Ok(response);
        //}

        //[HttpGet]
        //[Route("SplitDuration/GetBreakDownReasonLevel2")]
        //public async Task<IActionResult> GetBreakDownReasonLevel2(int LossCodeID)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.GetBreakDownReasonLevel2(LossCodeID);
        //    return Ok(response);
        //}

        //[HttpGet]
        //[Route("SplitDuration/GetBreakDownReasonLevel3")]
        //public async Task<IActionResult> GetBreakDownReasonLevel3(int LossCodeID)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.GetBreakDownReasonLevel3(LossCodeID);
        //    return Ok(response);
        //}

        //[HttpGet]
        //[Route("SplitDuration/DeleteSplitDuration")]
        //public async Task<IActionResult> DeleteSplitDuration(int tempModeId, int modeId)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.DeleteSplitDuration(tempModeId, modeId);
        //    return Ok(response);
        //}

        //[HttpGet]
        //[Route("SplitDuration/DeleteTempTableData")]
        //public async Task<IActionResult> DeleteTempTableData(int machineId, int modeId)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.DeleteTempTableData(machineId, modeId);
        //    return Ok(response);
        //}

        [HttpGet]
        [Route("SplitDuration/DeleteTempTableData")]
        public async Task<IActionResult> DeleteTempTableData(string tempmodeIds)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.DeleteTempTableData(tempmodeIds);
            return Ok(response);
        }

        [HttpPost]
        [Route("SplitDuration/SendMailToApproveOrReject")]
        public async Task<IActionResult> SendMailToApproveOrReject(SendMailDetails data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.SendMailToApproveOrReject(data);
            return Ok(response);
        }

        [HttpGet]
        [Route("SplitDuration/GetRejectReasons")]
        public async Task<IActionResult> GetRejectReasons()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetRejectReasons();
            return Ok(response);
        }

        [HttpPost]
        [Route("SplitDuration/SendMailAfterReject")]
        public async Task<IActionResult> SendMailAfterReject(SendMailDetails data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.SendMailAfterReject(data);
            return Ok(response);
        }

        /// <summary>
        /// AfterSplitDuration
        /// </summary>
        /// <param name="tempModeId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/AfterSplitDuration")]
        public async Task<IActionResult> AfterSplitDuration(int tempModeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.AfterSplitDuration(tempModeId);
            return Ok(response);
        }

        [HttpGet]
        [Route("SplitDuration/ApprovedSendToMainModeDetails")]
        public async Task<IActionResult> ApprovedSendToMainModeDetails(string tempModeIds)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.ApprovedSendToMainModeDetails(tempModeIds);
            return Ok(response);
        }

        [HttpPost]
        [Route("SplitDuration/GetApprovedDetails")]
        public async Task<IActionResult> GetApprovedDetails(SendMailDetails data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetApprovedDetails(data);
            return Ok(response);
        }

        //[HttpPost]
        //[Route("SplitDuration/GetModeDetailsTest")]
        //public async Task<IActionResult> GetModeDetailsTest(int MachineId, string date)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.GetModeDetailsTest(MachineId, date);
        //    return Ok(response);
        //}

        /// <summary>
        /// Index Page
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/Index")]
        public async Task<IActionResult> Index()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.Index();
            return Ok(response);
        }

        /// <summary>
        /// Add Individual Mode Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/AddIndividualModeDetails")]
        public async Task<IActionResult> AddIndividualModeDetails(AddIndividualMode data)
        {
            //calling DepartmentDAL busines layer
            CommonResponseWithTempId response = splitDuration.AddIndividualModeDetails(data);
            return Ok(response);
        }

        /// <summary>
        /// Get Reasons For Third Level
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetReasonsForThirdLevel")]
        public async Task<IActionResult> GetReasonsForThirdLevel(int tempModeId, int lossCodeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponseWithTempId response = splitDuration.GetReasonsForThirdLevel(tempModeId, lossCodeId);
            return Ok(response);
        }

        /// <summary>
        /// Get Time Details For Split Duration
        /// </summary>
        /// <param name="tempModeId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/GetTimeDetailsForSplitDuration")]
        public async Task<IActionResult> GetTimeDetailsForSplitDuration(int tempModeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.GetTimeDetailsForSplitDuration(tempModeId);
            return Ok(response);
        }

        /// <summary>
        /// Index
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/Index1")]
        public async Task<IActionResult> Index1()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.Index1();
            return Ok(response);
        }

        /// <summary>
        /// Split Duration
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/SplitDuration")]
        public async Task<IActionResult> SplitDuration(SplitDetails data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = splitDuration.SplitDuration(data);
            return Ok(response);
        }

        /// <summary>
        /// Delete TempMode Ids
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        //[HttpPost]
        //[Route("SplitDuration/DeleteTempModeIds")]
        //public async Task<IActionResult> DeleteTempModeIds(DeleteTempMode data)
        //{
        //    //calling DepartmentDAL busines layer
        //    CommonResponse1 response = splitDuration.DeleteTempModeIds(data);
        //    return Ok(response);
        //}

        /// <summary>
        /// Delete TempMode Ids
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/DeleteTempModeIds")]
        public async Task<IActionResult> DeleteTempModeIds(DeleteTempMode data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = splitDuration.DeleteTempModeIds(data);
            return Ok(response);
        }

        /// <summary>
        /// Update Split Duration
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("SplitDuration/UpdateSplitDuration")]
        public async Task<IActionResult> UpdateSplitDuration(UpdateTempMode data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = splitDuration.UpdateSplitDuration(data);
            return Ok(response);
        }

        /// <summary>
        /// Update Losses
        /// </summary>
        /// <param name="tempodeId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("SplitDuration/UpdateLosses")]
        public async Task<IActionResult> UpdateLosses(int tempodeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = splitDuration.UpdateLosses(tempodeId);
            return Ok(response);
        }
    }
}