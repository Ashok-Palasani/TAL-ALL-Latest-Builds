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
    //[Route("api/[controller]")]
    [ApiController]
    public class HMIScreenController : ControllerBase
    {
        IPlantShopCellData iPlantShopCellData;
        IHMIDetails iHMIDetails;

        public HMIScreenController(IPlantShopCellData _iPlantShopCellData, IHMIDetails _iHMIDetails)
        {
            iPlantShopCellData = _iPlantShopCellData;
            iHMIDetails = _iHMIDetails;
        }

        //Index Data
        [HttpGet]
        [Route("HMI/Index")]
        public async Task<ActionResult>Index()
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.Index();
            return Ok(obj);
        }


        //Get Plant Details
        [HttpGet]
        [Route("HMI/GetPlantDetails")]
        public async Task<ActionResult> GetPlant()
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetPlantDetails();
            return Ok(obj);
        }

        //Get Shop Details
        [HttpGet]
        [Route("HMI/GetShopDetails")]
        public async Task<ActionResult> GetShop(int plantId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetShopDetails(plantId);
            return Ok(obj);
        }

        //Get Cell Details
        [HttpGet]
        [Route("HMI/GetCellDetails")]
        public async Task<ActionResult> GetCell(int shopId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetCellDetails(shopId);
            return Ok(obj);
        }

        //Get Machine Details
        [HttpGet]
        [Route("HMI/GetMachineDetails")]
        public async Task<ActionResult> GetMachineDetails(int cellId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetMachineDetails(cellId);
            return Ok(obj);
        }


        [HttpPost]
        [Route("HMI/GetAllDetailsBasedOnPlantShopCell")]
        public async Task<ActionResult> GetAllDetailsBasedOnPlantShopCell(PlantShopCellGet data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetAllDetailsBasedOnPlantShopCell(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("HMI/GetAllDetailsBasedOnPlantShopCellMachine")]
        public async Task<ActionResult> GetAllDetailsBasedOnPlantShopCellMachine(PlantShopCellGet data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetAllDetailsBasedOnPlantShopCellMachine(data);
            return Ok(obj);
        }

        //To get plant,shop,cell,process details For Activity
        [HttpPost]
        [Route("HMI/GetAllDetailsBasedOnPlantShopCellProcess")]
        public async Task<ActionResult> GetAllDetailsBasedOnPlantShopCellProcess(PlantShopCellGet data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetAllDetailsBasedOnPlantShopCellProcess(data);
            return Ok(obj);
        }

        //Get Unassigned WorkOrder Details
        [HttpPost]
        [Route("HMI/GetMachineDetailsHMIData")]
        public async Task<ActionResult> GetMachineDetailsHMIData(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.GetMachineDetailsHMI(data);
            return Ok(obj);
        }

        //Get Login Details For Approvals
        [HttpPost]
        [Route("HMI/LoginDetails")]
        public async Task<ActionResult> LoginDet(LoginInfo data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = iHMIDetails.LoginDetails(data);
            return Ok(obj);
        }

        //Set the operatorId
        [HttpPost]
        [Route("HMI/SetOperator")]
        public async Task<ActionResult> SetOperatorID(OperatorDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SetOperatorId(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("HMI/GetDetailsAgain")]
        public async Task<ActionResult> GetdetailsAgain(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.DeletePreviousNoCode(data);
            return Ok(obj);
        }

        //Set the shift
        [HttpPost]
        [Route("HMI/SetShift")]
        public async Task<ActionResult> SetShiftID(SetShift data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SetShift(data);
            return Ok(obj);
        }

        //Split Duration
        [HttpGet]
        [Route("HMI/SplitDuration")]
        public async Task<ActionResult> SplitDuration(int uaWOId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SplitDuration(uaWOId);
            return Ok(obj);

        }

        //Edit the Current Data
        [HttpGet]
        [Route("HMI/UpdateData")]
        public async Task<ActionResult> UpdateData(int uaWOId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.Edit(uaWOId);
            return Ok(obj);

        }

        //Delete  duration
        [HttpGet]
        [Route("HMI/DeleteDuration")]
        public async Task<ActionResult> DeleteDuration(int uaWOId, string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.DeleteSplitDuration(uaWOId, uaWOIds);
            return Ok(obj);
        }

        //validate and store the split duration
        [HttpPost]
        [Route("HMI/ValidateDuration")]
        public async Task<ActionResult> ValidateDuration(CompareDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.ComapreEndDuration(data);
            return Ok(obj);
        }

        //Clear the split duration
        [HttpGet]
        [Route("HMI/ClearDuration")]
        public async Task<ActionResult>ClearDuration(string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.ClearAllData(uaWOIds);
            return Ok(obj);
        }

        //confirm the splitduration
        [HttpPost]
        [Route("HMI/ConfirmSplitDuration")]
        public async Task<ActionResult> ConfirmSplitDurationData(ConfirmSplitDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.ConfirmSpliDuration(data);
            return Ok(obj);
        }


        #region Old logic       


        ////Split Duration
        //[HttpGet]
        //[Route("HMI/SplitDuration")]
        //public async Task<ActionResult>SplitDuration(int uaWOId)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj.response = iHMIDetails.SplitDuration(uaWOId);
        //    return Ok(obj);

        //}

        ////validate and store the duration
        //[HttpPost]
        //[Route("HMI/ValidateDuration")]
        //public async Task<ActionResult>ValidateDuration(CompareDuration data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.ComapreEndDuration(data);
        //    return Ok(obj);
        //}

        #endregion

        ////job finish the wororder
        [HttpPost]
        [Route("HMI/JobFinishWorkOrder")]
        public async Task<ActionResult> JobFinishWO(StoreHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.JobFinishWorkOrderDetails(data);
            return Ok(obj);
        }


        ////Set the row with blank
        [HttpGet]
        [Route("HMI/RowBlank")]
        public async Task<ActionResult> RowBlank(int unwoId, int value)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SetBlank(unwoId, value);
            return Ok(obj);
        }

        //partial finish the wororder

        [HttpPost]
        [Route("HMI/PartialFinishWorkOrder")]
        public async Task<ActionResult> PartialFinishWO(StoreHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.PartialFinishWorkOrderDetails(data);
            return Ok(obj);
        }

        //Remove UnAssigned Work order
        [HttpGet]
        [Route("HMI/RemoveWO")]
        public async Task<ActionResult> RemoveUnAssignedWO(int uawoid)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.RemoveWorkOrder(uawoid);
            return Ok(obj);
        }


        // Get DDL list full and work center wise
        [HttpPost]
        [Route("HMI/DDLList")]
        public async Task<ActionResult> GetDDLList(DDLList data)
        {
            DDLCommonResponse obj = new DDLCommonResponse();
            obj = iHMIDetails.GetDDLList(data);
            return Ok(obj);
        }

        //Validate DDL selection
        [HttpPost]
        [Route("HMI/DLLValidate")]
        public async Task<ActionResult> ValidateDDL(DDLIds data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.ValidateDDLIDS(data);
            return Ok(obj);
        }
                     

        //Set the reworkorder
        [HttpPost]
        [Route("HMI/SetReWork")]
        public async Task<ActionResult> SetReWorkOrder(SetReWork data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SetReWorkOrder(data);
            return Ok(obj);
        }

        // Get shop wise work center
        [HttpGet]
        [Route("HMI/ShopWiseWorkCenter")]
        public async Task<ActionResult> GetShopWiseWorkCenter(int uaWoID)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.GetShopWiseWorkCenter(uaWoID);
            return Ok(obj);
        }

        //Send the Work Order details from DDL
        [HttpPost]
        [Route("HMI/SendWorkOrderDet")]
        public async Task<ActionResult> SendWorkOrderDetails(SendDDLUnAsignedWoId data)
        {
            CommonResponse obj = new CommonResponse();
            List<SelectWO> listSelectWO = new List<SelectWO>();
            int unAsignedWOId = data.unAsignedId;
            string[] splitDDLIds = data.ddlIds.Split(',');
            for (int i = 0; i < splitDDLIds.Count(); i++)
            {
                int ddlId = Convert.ToInt32(splitDDLIds[i]);
                bool check = iHMIDetails.CheckPrvDDL(ddlId);
                if (check)
                {
                    SelectWO objSelectWO = new SelectWO();
                    objSelectWO.uaWOId = unAsignedWOId;
                    objSelectWO.ddlId = splitDDLIds[i];
                    listSelectWO.Add(objSelectWO);
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "The WorkOrder Is Already Selected for Previous UNAssigned WorkOrder Timings";
                    break;
                }
            }
            if (listSelectWO.Count > 0)
            {
                obj = iHMIDetails.SendWorkOrders(listSelectWO);
            }           
            return Ok(obj);
        }

        ////Start WO
        //[HttpGet]
        //[Route("HMI/StartWO")]
        //public async Task<ActionResult> StartWorkOrder(int uaWoID)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.StartWO(uaWoID);
        //    return Ok(obj);
        //}


        ////Get the Hold Code
        //[HttpPost]
        //[Route("HMI/GetHoldCode")]
        //public async Task<ActionResult> GetHoldCodeReason(HoldCode data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.GetHoldCodes(data);
        //    return Ok(obj);
        //}

        ////Set the Hold Code
        //[HttpPost]
        //[Route("HMI/SetHoldCode")]
        //public async Task<ActionResult> SetHoldCodeReason(HoldCodeGet data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.SetHoldCodereason(data);
        //    return Ok(obj);
        //}

        ////Get the Generic Code
        //[HttpPost]
        //[Route("HMI/GetGenericCode")]
        //public async Task<ActionResult> GetGenericCodeReason(GenericCode data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.GetGenericCodes(data);
        //    return Ok(obj);
        //}

        ////Set the Generic Code
        //[HttpPost]
        //[Route("HMI/SetGenericCode")]
        //public async Task<ActionResult> SetGenericCodeReason(GenericCodeGet data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.SetGenericCodereason(data);
        //    return Ok(obj);
        //}

        //Set the Split WO
        [HttpPost]
        [Route("HMI/SetSplitWO")]
        public async Task<ActionResult> SetSplitWorkOrder(SetSplitWork data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SplitWorkOrder(data);
            return Ok(obj);
        }

        // Send To Approve
        [HttpPost]
        [Route("HMI/SendToApproveWO")]
        public async Task<ActionResult> SendToApproveAllWO(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.SendToApproveAllWODetails(data);
            return Ok(obj);
        }

        // Accept The WO
        [HttpPost]
        [Route("HMI/AcceptWO")]
        public async Task<ActionResult> AcceptAllWO(EntityHMIDetails data)
        {
            //iHMIDetails.TakeBackupReportData(correctedDate);
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.AcceptAllWODetails(data);
            return Ok(obj);
        }

        // Get Reject Reasons
        [HttpGet]
        [Route("HMI/GetRejectReason")]
        public async Task<ActionResult> GetRejectReassons()
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.GetRejectReason();
            return Ok(obj);
        }

        // Reject The WO
        [HttpPost]
        [Route("HMI/RejectWO")]
        public async Task<ActionResult> RejectAllWO(RejectReasonStore data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.RejectAllWODetails(data);
            return Ok(obj);
        }

        //Accept Reject Unassigned wo table
        [HttpPost]
        [Route("HMI/AcceptRejectUNWOTable")]
        public async Task<ActionResult> AcceptRejectUaAssignedWOTable(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.AcceptRejectUnWOTable(data);
            return Ok(obj);
        }

        /// <summary>
        /// Un assigned WO Split Duration Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("HMI/UnassignedWOSplitDurationDetails")]
        public async Task<IActionResult> UnassignedWOSplitDurationDetails(List<UnassignedWOStartEndDateTime> data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = iHMIDetails.UnassignedWOSplitDurationDetails(data);
            return Ok(response);
        }

        //when click on update button update the endtime
        [HttpPost]
        [Route("HMI/UpdateEndTime")]
        public async Task<ActionResult> EditEndTime(CompareDuration data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIDetails.UpdateTime(data);
            return Ok(obj);
        }

        ////On Generic workorder
        //[HttpPost]
        //[Route("HMI/GenericWorkOrder")]
        //public async Task<ActionResult> GenericWorkOrder(GenericWO genericWO)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.GenericWorkOrder(genericWO);
        //    return Ok(obj);
        //}


        // Update to report table
        //[HttpPost]
        //[Route("HMI/UpdateToreportTable")]
        //public async Task<ActionResult> UpdateLiveHmiTab()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iHMIDetails.UpdateToReportTables();
        //    return Ok(obj);
        //}
    }
}
