using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.DAL;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using static DAS.EntityModels.OPCancelEntity;
using System.Net.Http;
using System.Net;
using static DAS.EntityModels.SplitDurationEntity;
//using System.Web.Http;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class PrevOpCancelController : ControllerBase
    {
        IOpCancel _iOpCancel;

        public PrevOpCancelController(IOpCancel iOpCancel)
        {
            _iOpCancel = iOpCancel;
        }

        //Index
        [HttpGet]
        [Route("OPCan/Index")]
        public async Task<IActionResult> Index()
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.Index();
            return Ok(obj);
        }


        //Get Plant Details
        [HttpPost]
        [Route("OPCan/PrevOpCancelDetails")]
        public async Task<IActionResult> PrevOpCancelDetails(LsitOPcancelDet PrevOpCancelDetails)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
               //var list =JsonConvert.DeserializeObject(PrevOpCancelDetails);


                List<OPCancelDetails> OPcancelDet = new List<OPCancelDetails>();


                obj = _iOpCancel.GetPrevOpCancelDetails(PrevOpCancelDetails);
            }
            catch(Exception ex)
            {
                
            }

            return Ok(obj);
        }


        //Get the excel data and insert and show case in view
        [HttpPost]
        [Route("OPCan/GetPreiousOPData")]
        public async Task<ActionResult> GetPreiousOperationData(List<UpLoadExcel> data)
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.GetListOFOperationNumberFileData(data);
            return Ok(obj);
        }

        //select and unselect the operatio number
        [HttpPost]
        [Route("OPCan/AcceptRejectCheckBox")]
        public async Task<ActionResult> AcceptRejectOperation(AcceptCancel data)
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.AcceptRejectOperationNo(data);
            return Ok(obj);
        }

        //Get All unchecked Data
        [HttpGet]
        [Route("OPCan/GetCancledData")]
        public async Task<ActionResult> GetCancledData()
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.GetCancledList();
            return Ok(obj);
        }

        //Get All checked Data
        [HttpGet]
        [Route("OPCan/GetUnCancledData")]
        public async Task<ActionResult> GetUnCancledData()
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.GetUnCancledList();
            return Ok(obj);
        }

        //Send Approval For accept and reject
        [HttpGet]
        [Route("OPCan/SendApproval")]
        public async Task<ActionResult>SendForApproval()
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.SendForApproval();
            return Ok(obj);
        }

        //Get Reassons for Opertion cancellation
        [HttpGet]
        [Route("OPCan/GetRejectReason")]
        public async Task<ActionResult>GetRejectReason()
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.GetRejectReason();
            return Ok(obj);
        }

        //Accept the Operation cancallation
        [HttpPost]
        [Route("OPCan/AcceptPrvOp")]
        public async Task<ActionResult> AcceptPrvOp(ApprovalClass data)
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.AcceptPrvOp(data);
            return Ok(obj);
        }

        //To get Login Details For Approval
        [HttpPost]
        [Route("OPCan/LoginDetails")]
        public async Task<ActionResult> Login(LoginInfo data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = _iOpCancel.LoginDetails(data);
            return Ok(obj);
        }

        //Reject the Operation cancallation
        [HttpPost]
        [Route("OPCan/RejectPrvOp")]
        public async Task<ActionResult> RejectPrvOp(RejectClass data)
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.RejectPrvOPCan(data);
            return Ok(obj);
        }

        //Accept reject common table
        [HttpPost]
        [Route("OPCan/CommonData")]
        public async Task<ActionResult> CommonData(ApprovalClass data)
        {
            CommonResponse obj = new CommonResponse();
            obj = _iOpCancel.CommonDtatForDisplay(data);
            return Ok(obj);
        }

    }
}