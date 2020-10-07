using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.CommonResponseWithMachineName;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Controllers
{
    [ApiController]
    public class TCFController : ControllerBase
    {
        INoCodeInterface tcf;
        IPlantShopCellData iPlantShopCellData;
        public TCFController(INoCodeInterface tcfcontroller, IPlantShopCellData _iPlantShopCellData)
        {
            tcf = tcfcontroller;
            iPlantShopCellData = _iPlantShopCellData;
        }

        //Index Get Data
        [HttpGet]
        [Route("LossCode/Index")]
        public async Task<ActionResult>Index()
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.Index();
            return Ok(obj);
        }

        //Get Plant Details
        [HttpGet]
        [Route("LossCode/GetPlantDetails")]
        public async Task<ActionResult> GetPlant()
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetPlantDetails();
            return Ok(obj);
        }

        //Get Shop Details
        [HttpGet]
        [Route("LossCode/GetShopDetails")]
        public async Task<ActionResult> GetShop(int plantId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetShopDetails(plantId);
            return Ok(obj);
        }

        //Get Cell Details
        [HttpGet]
        [Route("LossCode/GetCellDetails")]
        public async Task<ActionResult> GetCell(int shopId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetCellDetails(shopId);
            return Ok(obj);
        }

        //Get Machine Details
        [HttpGet]
        [Route("LossCode/GetMachineDetails")]
        public async Task<ActionResult> GetMachineDetails(int cellId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetMachineDetails(cellId);
            return Ok(obj);
        }
        [Route("LossCode/GetLoss")]
        [HttpPost]
        public async Task<ActionResult> GetLossDet(lossmodel obj)
        {
            EntityModelWithLossCount entity = new EntityModelWithLossCount();
            entity = tcf.GetLossDet(obj);
            return Ok(entity);
        }

        #region split duration

        //Split Duration
        [HttpGet]
        [Route("LossCode/SplitDuration")]
        public async Task<ActionResult> SplitDuration(int uaWOId)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.SplitDuration(uaWOId);
            return Ok(obj);

        }

        //Get Login Details For Approvals
        [HttpPost]
        [Route("LossCode/LoginDetails")]
        public async Task<ActionResult> GetLogin(LoginInfo data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = tcf.LoginDetails(data);
            return Ok(obj);

        }

        //Delete  duration
        [HttpGet]
        [Route("LossCode/DeleteDuration")]
        public async Task<ActionResult> DeleteDuration(int uaWOId, string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.DeleteSplitDuration(uaWOId, uaWOIds);
            return Ok(obj);
        }

        //validate and store the split duration
        [HttpPost]
        [Route("LossCode/ValidateDuration")]
        public async Task<ActionResult> ValidateDuration(CompareDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.ComapreEndDuration(data);
            return Ok(obj);
        }

        //Clear the split duration
        [HttpGet]
        [Route("LossCode/ClearDuration")]
        public async Task<ActionResult> ClearDuration(string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.ClearAllData(uaWOIds);
            return Ok(obj);
        }

        //confirm the splitduration
        [HttpPost]
        [Route("LossCode/ConfirmSplitDuration")]
        public async Task<ActionResult> ConfirmSplitDurationData(ConfirmSplitDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.ConfirmSpliDuration(data);
            return Ok(obj);
        }

        #endregion



        [Route("LossCode/DeletePreviousRecord")]
        [HttpPost]
        public async Task<ActionResult> DeleteRecord(lossmodel obj)
        {
            CommonResponse entity = new CommonResponse();
            entity = tcf.DeletePreviousNoCode(obj);
            return Ok(entity);
        }

        [Route("LossCode/GetLossCodeLevel1")]
        [HttpGet]
        public async Task<ActionResult> GetLossCodeLevel1()
        {
            EntityModel entity = new EntityModel();
            entity = tcf.GetLossCodeLevel1();
            return Ok(entity);
        }
        [Route("LossCode/GetLossCodeLevel2")]
        [HttpPost]
        public async Task<ActionResult> GetLossCodeLevel2(int LossCodeID)
        {
            EntityModel entity = new EntityModel();
            entity = tcf.GetLossCodeLevel2(LossCodeID);
            return Ok(entity);
        }
        [Route("LossCode/GetLossCodeLevel3")]
        [HttpPost]
        public async Task<ActionResult> GetLossCodeLevel3(int LossCodeID)
        {
            EntityModel entity = new EntityModel();
            entity = tcf.GetLossCodeLevel3(LossCodeID);
            return Ok(entity);
        }
        [Route("LossCode/Updatetcfloss")]
        [HttpPost]
        public async Task<ActionResult> updatetcfrecord(updateLoss obj)
        {
            EntityModel entity = new EntityModel();
            entity = tcf.UpdatetcfLoss(obj);
            return Ok(entity);
        }
        [Route("LossCode/SendToApproveAllLossDetails")]
        [HttpPost]
        public async Task<ActionResult> GettcfLoss(EntityHMIDetails data)
        {
            EntityModel entity = new EntityModel();
            entity = tcf.SendToApproveAllLossDetails(data);
            return Ok(entity);
        }
        //[Route("LossCode/GettcfLossLevel")]
        //[HttpGet]
        //public async Task<ActionResult> GettcfLossLevel(int LossID)
        //{
        //    EntityModel entity = new EntityModel();
        //    entity = tcf.GettcfLossLevel(LossID);
        //    return Ok(entity);
        //}
        [Route("LossCode/UpdateliveLoss")]
        [HttpPost]
        public async Task<ActionResult> UpdateliveLoss(EntityHMIDetails data)
        {
            //EntityModel entity = new EntityModel();
            CommonResponse comobj = new CommonResponse();
            //entity = tcf.UpdateliveLoss(data.fromDate);
            comobj=tcf.AcceptAllNoCodeDetails(data);
            return Ok(comobj);
        }
        //[Route("LossCode/InsertMultiLoss")]
        //[HttpPost]
        //public async Task<ActionResult> InsertMultiLoss(tcfclass obj, int noofreason)
        //{
        //    EntityModel entity = new EntityModel();
        //    entity = tcf.InsertMultiLoss(obj, noofreason);
        //    return Ok(entity);
        //}
        [Route("LossCode/SetRejectReason")]
        [HttpPost]
        public async Task<ActionResult> SetRejecttcfLoss(RejectReasonStore data)
        {
            CommonResponse entity = new CommonResponse();
            entity = tcf.RejectAllNoCodeDetails(data);
            return Ok(entity);
        }

        [Route("LossCode/GetAcceptTaleList")]
        [HttpPost]
        public async Task<ActionResult> GetAccept(EntityHMIDetails data)
        {
            CommonResponse entity = new CommonResponse();
            entity = tcf.AcceptRejectNoCodeTable(data);
            return Ok(entity);
        }

        [Route("LossCode/GetPlantShopCellMachineNames")]
        [HttpPost]
        public async Task<ActionResult> GetPlantShopCellMachineNames(PlantShopCellMachineId data)
        {
            CommonResponse entity = new CommonResponse();
            entity = iPlantShopCellData.GetPlantShopCellMachineNames(data);
            return Ok(entity);
        }

        [Route("LossCode/GetRejectReason")]
        [HttpGet]
        public async Task<ActionResult> GetReason()
        {
            EntityModel entity = new EntityModel();
            entity = tcf.GetReason();
            return Ok(entity);
        }

        [HttpPost]
        [Route("HMI/NoCodeSplitDurationDetails")]
        public async Task<IActionResult> NoCodeSplitDurationDetails(List<NoCodeStartEndDateTime> data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcf.NoCodeSplitDurationDetails(data);
            return Ok(response);
        }

        //Update the Previous EndTime
        [HttpPost]
        [Route("LossCode/UpdateEndTime")]
        public async Task<ActionResult> UpdateEndTime(CompareDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = tcf.UpdateTime(data);
            return Ok(obj);
        }
    }
}
