using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Mvc;


namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class ActivityController : Controller
    {
        IActivity actobj;

        public ActivityController(IActivity act)
        {
            actobj = act;
        }

        //Get Activity Details
        [HttpGet]
        [Route("Activity/GetActivityDetails")]
        public async Task<ActionResult> GetAct()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetActivity();
            return Ok(obj);
        }

        //Create New Activity
        [HttpPost]
        [Route("Activity/CreateActivityDetails")]
        public async Task<ActionResult> CreateActivity(EntityActivity data)
        {
            EntityModel obj = new EntityModel();
          obj = actobj.CreateActivity(data);
            return Ok(obj);
        }

        //Get Individual Activity to Edit
        //[HttpGet]
        //[Route("Activity/EditActivityDetails")]
        //public async Task<ActionResult> EditActivity(int id)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.EditActivityDet(id);
        //    return Ok(obj);
        //}

        //Update Existing Activity
        [HttpPost]
        [Route("Activity/UpdateActivityDetails")]
        public async Task<ActionResult> UpdateActivity(EntityActivity data)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.UpdateActivity(data);
            return Ok(obj);
        }

        //Update Existing Activity
        [HttpGet]
        [Route("Activity/DeleteActivityDetails")]
        public async Task<ActionResult>DeleteActivity(int id)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.DeleteActivity(id);
            return Ok(obj);
        }

        //Get Process List
        [HttpGet]
        [Route("Activity/ProcessList")]
        public async Task<ActionResult> ProcessList()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetProcessList();
            return Ok(obj);
        }
        //Validating the Process whether its already exists for that perticular cell
        [HttpGet]
        [Route("Activity/ValidateProcess")]
        public async Task<ActionResult> validateProcessDet(int cellId,int processId)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.validateprocess(cellId,processId);
            return Ok(obj);
        }

        //Validating the Activity whether its already exists for that perticular cell and process
        [HttpPost]
        [Route("Activity/ValidateActivity")]
        public async Task<ActionResult> validateActivityDet(validateprocessCell data)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.validateActivity(data);
            return Ok(obj);
        }

        //To Upload Activity Excel
        //[HttpPost]
        //[Route("Activity/AddUploadedActivityDetails")]
        //public async Task<ActionResult> AddUploadedActivityDetails(actlist data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.AddUploadedActivityDetails(data);
        //    return Ok(obj);
        //}
    }
}
