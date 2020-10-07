using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class NoLoginController : ControllerBase
    {
        IPlantShopCellData iPlantShopCellData;
        INoLogin iNoLogin;

        public NoLoginController(IPlantShopCellData _iPlantShopCellData, INoLogin _iNoLogin)
        {
            iPlantShopCellData = _iPlantShopCellData;
            iNoLogin = _iNoLogin;
        }


        /// <summary>
        /// Split Duration Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("NoLogin/SplitDurationDetails")]
        public async Task<IActionResult> SplitDurationDetails(List<NoLoginStartEndDateTime> data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = iNoLogin.SplitDurationDetails(data);
            return Ok(response);
        }

        [HttpPost]
        [Route("NoLogin/SplitDuration")]
        public async Task<IActionResult> SplitDuration(NoSplitDetails data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = iNoLogin.SplitDuration(data);
            return Ok(response);
        }


        /// <summary>
        /// DeleteTempTableData
        /// </summary>
        /// <param name="noLoginId"></param>
        /// <param name="machineId"></param>
        /// <returns></returns>
        //[HttpGet]
        //[Route("NoLogin/DeleteTempTableData")]
        //public async Task<ActionResult> DeleteTempTableData(int noLoginId, int machineId)
        //{
        //    CommonResponse1 response = iNoLogin.DeleteTempTableData(noLoginId, machineId);
        //    return Ok(response);
        //}

        /// <summary>
        /// DeleteTempTableData
        /// </summary>
        /// <param name="noLoginId"></param>
        /// <param name="machineId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("NoLogin/DeleteTempTableData")]
        public async Task<ActionResult> DeleteTempTableData(string noLoginIds)
        {
            CommonResponse1 response = iNoLogin.DeleteTempTableData(noLoginIds);
            return Ok(response);
        }


        //Index To get data
        [HttpGet]
        [Route("Index")]
        public async Task<ActionResult>Index()
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.Index();
            return Ok(obj);
        }


        //Get Plant Details
        [HttpGet]
        [Route("GetPlantDetails")]
        public async Task<ActionResult> GetPlant()
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetPlantDetails();
            return Ok(obj);
        }

        /// <summary>
        /// GetTimeDetails
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("NoLogin/GetTimeDetails")]
        public async Task<IActionResult> GetTimeDetails(int id)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = iNoLogin.GetTimeDetails(id);
            return Ok(response);
        }

        [HttpPost]
        [Route("NoLogin/DeleteSplitDuration")]
        public async Task<IActionResult> DeleteSplitDuration(DeleteSplitDuration data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = iNoLogin.DeleteSplitDuration(data);
            return Ok(response);
        }
        //To Get Login Details
        [HttpPost]
        [Route("NoLogin/LoginDetails")]
        public async Task<IActionResult> GetLogindet(LoginInfo data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = iNoLogin.LoginDetails(data);
            return Ok(response);
        }

        /// <summary>
        /// After Split Duration
        /// </summary>
        /// <param name="date"></param>
        /// <param name="machineId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("NoLogin/AfterSplitDuration")]
        public async Task<ActionResult> AfterSplitDuration(string date, int machineId)
        {
            CommonResponse1 response = iNoLogin.AfterSplitDuration(date, machineId);
            return Ok(response);
        }

        [HttpPost]
        [Route("NoLogin/UpdateSplitDuration")]
        public async Task<IActionResult> UpdateSplitDuration(UpdateNoLogin data)
        {
            //calling DepartmentDAL busines layer
            GeneralResponse1 response = iNoLogin.UpdateSplitDuration(data);
            return Ok(response);
        }

        /// <summary>
        /// DeleteTempTableData
        /// </summary>
        /// <param name="noLoginId"></param>
        /// <param name="machineId"></param>
        /// <returns></returns>
        //[HttpGet]
        //[Route("NoLogin/DeleteTempTableData")]
        //public async Task<ActionResult> DeleteTempTableData(int noLoginId, int machineId)
        //{
        //    CommonResponse1 response = iNoLogin.DeleteTempTableData(noLoginId, machineId);
        //    return Ok(response);
        //}

        //Get Shop Details
        [HttpGet]
        [Route("GetShopDetails")]
        public async Task<ActionResult> GetShop(int plantId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetShopDetails(plantId);
            return Ok(obj);
        }

        //Get Cell Details
        [HttpGet]
        [Route("GetCellDetails")]
        public async Task<ActionResult> GetCell(int shopId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetCellDetails(shopId);
            return Ok(obj);
        }

        //Get Machine Details
        [HttpGet]
        [Route("GetMachineDetails")]
        public async Task<ActionResult> GetMachineDetails(int cellId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetMachineDetails(cellId);
            return Ok(obj);
        }

        //Getting the Nologin timing data to index
        [HttpPost]
        [Route("GettingTimeForNoLogin")]
        public async Task<ActionResult> GettingTimeForNoLogin(EntityNoLogin data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = iNoLogin.GetNoLoginTiming(data);
            return Ok(obj);
        }
        
        [HttpPost]
        [Route("GettingTimeForNoLoginagain")]
        public async Task<ActionResult> GettingTimeForNoLoginagain(EntityNoLogin data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.DeletePreviousNoCode(data);
            return Ok(obj);
        }

        //Set the operatorId
        [HttpPost]
        [Route("SetOperator")]
        public async Task<ActionResult> SetOperatorID(NoLoginOperatorDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SetOperatorId(data);
            return Ok(obj);
        }

        //Set the shift
        [HttpPost]
        [Route("SetShift")]
        public async Task<ActionResult> SetShiftID(NoLoginSetShift data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SetShift(data);
            return Ok(obj);
        }

        // Get shop wise work center
        [HttpGet]
        [Route("ShopWiseWorkCenter")]
        public async Task<ActionResult> GetShopWiseWorkCenter(int noLoginId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.GetShopWiseWorkCenter(noLoginId);
            return Ok(obj);
        }


        // Edit the Current data
        [HttpGet]
        [Route("UpdateRecord")]
        public async Task<ActionResult> UpdateRecord(int noLoginId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.Edit(noLoginId);
            return Ok(obj);
        }

        // Get DDL list full and work center wise
        [HttpPost]
        [Route("DDLList")]
        public async Task<ActionResult> GetDDLList(NoLoginDDLList data)
        {
            NoLoginDDLCommonResponse obj = new NoLoginDDLCommonResponse();
            obj = iNoLogin.GetDDLList(data);
            return Ok(obj);
        }

        //Validate DDL selection
        [HttpPost]
        [Route("DLLValidate")]
        public async Task<ActionResult> ValidateDDL(DDLIds data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.ValidateDDLIDS(data);
            return Ok(obj);
        }

        //Send the Work Order details from DDL
        [HttpPost]
        [Route("SendWorkOrderDet")]
        public async Task<ActionResult> SendWorkOrderDetails(NoLoginSendDDLUnAsignedWoId data)
        {
            CommonResponse obj = new CommonResponse();
            List<NoLoginSelectWO> listSelectWO = new List<NoLoginSelectWO>();
            int unAsignedWOId = data.noLoginId;
            string[] splitDDLIds = data.ddlIds.Split(',');
            for (int i = 0; i < splitDDLIds.Count(); i++)
            {
                int ddlId = Convert.ToInt32(splitDDLIds[i]);
                bool check = iNoLogin.CheckPrvDDL(ddlId);
                if (check)
                {
                    NoLoginSelectWO objSelectWO = new NoLoginSelectWO();
                    objSelectWO.noLoginId = unAsignedWOId;
                    objSelectWO.ddlId = splitDDLIds[i];
                    listSelectWO.Add(objSelectWO);
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "The WorkOrder Is Already Selected for Previous NoLogin WorkOrder Timings";
                    break;
                }
            }
            if (listSelectWO.Count > 0)
            {
                obj = iNoLogin.SendWorkOrders(listSelectWO);
            }
            return Ok(obj);
        }

        //Remove UnAssigned Work order
        [HttpGet]
        [Route("RemoveWO")]
        public async Task<ActionResult> RemoveUnAssignedWO(int noLoginId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.RemoveWorkOrder(noLoginId);
            return Ok(obj);
        }

        //Set the Split WO
        [HttpPost]
        [Route("SetSplitWO")]
        public async Task<ActionResult> SetSplitWorkOrder(NoLoginSetSplitWork data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SplitWorkOrder(data);
            return Ok(obj);
        }

        //Set the reworkorder
        [HttpPost]
        [Route("SetReWork")]
        public async Task<ActionResult> SetReWorkOrder(NoLoginSetReWork data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SetReWorkOrder(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("PartialFinishWorkOrder")]
        public async Task<ActionResult> PartialFinishWO(NoLoginStoreHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.PartialFinishWorkOrderDetails(data);
            return Ok(obj);
        }


        ////job finish the wororder
        [HttpPost]
        [Route("JobFinishWorkOrder")]
        public async Task<ActionResult> JobFinishWO(NoLoginJobFinish data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.JobFinishWorkOrderDetails(data);
            return Ok(obj);
        }

        ////Set the row with blank
        [HttpGet]
        [Route("RowBlank")]
        public async Task<ActionResult> RowBlank(int noLoginId, int value)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SetBlank(noLoginId, value);
            return Ok(obj);
        }


        // Send To Approve
        [HttpPost]
        [Route("SendToApproveWO")]
        public async Task<ActionResult> SendToApproveAllWO(EntityNoLogin data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.SendToApproveAllWODetails(data);
            return Ok(obj);
        }

        // Accept The WO
        [HttpPost]
        [Route("HMI/AcceptWO")]
        public async Task<ActionResult> AcceptAllWO(EntityNoLogin data)
        {
            //iHMIDetails.TakeBackupReportData(correctedDate);
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.AcceptAllWODetails(data);
            return Ok(obj);
        }

        // Get Reject Reasons
        [HttpGet]
        [Route("HMI/GetRejectReason")]
        public async Task<ActionResult> GetRejectReassons()
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.GetRejectReason();
            return Ok(obj);
        }

        // Reject The WO
        [HttpPost]
        [Route("HMI/RejectWO")]
        public async Task<ActionResult> RejectAllWO(RejectReasonStore data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.RejectAllWODetails(data);
            return Ok(obj);
        }

        //Accept Reject Unassigned wo table
        [HttpPost]
        [Route("HMI/AcceptRejectUNWOTable")]
        public async Task<ActionResult> AcceptRejectUaAssignedWOTable(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iNoLogin.AcceptRejectUnWOTable(data);
            return Ok(obj);
        }  
        
    }
}