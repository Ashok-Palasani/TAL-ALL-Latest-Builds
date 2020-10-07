using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.SplitDurationEntity;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class HMIWrongQtyController : Controller
    {
        IHMIWrongQty iHMIWrongQty;
        public HMIWrongQtyController(IHMIWrongQty _iHMIWrongQty)
        {
            iHMIWrongQty = _iHMIWrongQty;
        }

        #region Auto Suggest


        //get work order details
        [HttpGet]
        [Route("HMIWrongQty/GetWorkOrderDetails")]
        public async Task<ActionResult> GetWorkOrderDetails(string workOrderNo)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.GetWorkOrderDetails(workOrderNo);
            return Ok(obj);
        }


        //get Operation details
        [HttpGet]
        [Route("HMIWrongQty/GetOperationDetails")]
        public async Task<ActionResult> GetOperationDetails(string workOrderNo, string operationNo)
        {
            CommonResponse obj = new CommonResponse();
            // due to anguldar chnges we made it like this
            AutoSuggestOpetaion operationData = new AutoSuggestOpetaion();
            operationData.workOrderNo = workOrderNo;
            operationData.operationNo = operationNo;
            obj = iHMIWrongQty.GetOperationDetails(operationData);
            return Ok(obj);
        }

        #endregion

        //index method
        [HttpGet]
        [Route("HMIWrongQty/Index")]
        public async Task<ActionResult>Index()
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.IndexWOQty();
            return Ok(obj);
        }


        //Get the JF and PF Details
        [HttpPost]
        [Route("HMIWrongQty/GetJFOrPFDetails")]
        public async Task<ActionResult> GetWOJFPFDetails(Getdata data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse obj = new CommonResponse();
            obj=iHMIWrongQty.GetWOJFPFDetails(data);
            return Ok(obj);
        }

        //update the qty and validate   lodic for chnaging the jf and pf values too
        [HttpPost]
        [Route("HMIWrongQty/UpateWQTY")]
        public async Task<ActionResult> UpateWQTY(ValidateQTYUpdate data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse obj = new CommonResponse();
            obj=  iHMIWrongQty.ValidateQtyUpdate(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("HMIWrongQty/ValidateTheQtyIds")]
        public async Task<ActionResult> ValidateTheQtyIds(GetIdsValues data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.ValidateQtyData(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("HMIWrongQty/PartialFinish")]
        public async Task<ActionResult> pf(pf data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = iHMIWrongQty.PartialFinish(data);
            return Ok(obj);
        }

        [HttpPost]
        [Route("HMIWrongQty/JobFinish")]
        public async Task<ActionResult> jf(pf data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            obj = iHMIWrongQty.JobFinish(data);
            return Ok(obj);
        }


        //Send for approval
        [HttpPost]
        [Route("HMIWrongQty/SendMail")]          
        public async Task<ActionResult>SendMailAll(Getdata data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.SendApproval(data);
            return Ok(obj);
        }

        // common method for getting the tbale structer to display aceept and reject data
        [HttpPost]
        [Route("HMIWrongQty/GetAllData")]
        public async Task<ActionResult>GetAllWOQtyData(GetMaildata data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.GetAllData(data);
            return Ok(obj);
        }

        //accept the data sent for approval
        [HttpPost]
        [Route("HMIWrongQty/AcceptData")]
        public async Task<ActionResult>AcceptAllWOQtyData(Getdata data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.AcceptWoQtyData(data);
            return Ok(obj);
        }

        //To Get Login Details
        [HttpPost]
        [Route("NoLogin/LoginDetails")]
        public async Task<IActionResult> GetLogindet(LoginInfo data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = iHMIWrongQty.LoginDetails(data);
            return Ok(response);
        }

        //get the reject reasons
        [HttpGet]
        [Route("HMIWrongQty/GetReasons")]
        public async Task<ActionResult> GetReasons()
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.GetRejectReasons();
            return Ok(obj);
        }


        // reject the data by reject reason
        [HttpPost]
        [Route("HMIWrongQty/RejectData")]
        public async Task<ActionResult> RejectAllWOQtyData(GetdataWithRejectReason data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iHMIWrongQty.RejectWoQtyData(data);
            return Ok(obj);
        }

    }
}
