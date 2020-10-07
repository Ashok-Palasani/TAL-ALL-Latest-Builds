using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class ProcessController : Controller
    {
        IProcess actobj;

        public ProcessController(IProcess act)
        {
            actobj = act;
        }

        //Get Activity Details
        [HttpGet]
        [Route("Activity/GetProcessDetails")]
        public async Task<ActionResult> GetProcess()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetProcess();
            return Ok(obj);
        }

        //Create New Activity
        [HttpPost]
        [Route("Activity/CreateOrEditProcessDetails")]
        public async Task<ActionResult> CreateProcess(EntityProcess data)
        {
            EntityModel obj = new EntityModel();
            obj = actobj.CreateProcess(data);
            return Ok(obj);
        }

        //Get Individual Activity to Edit
        //[HttpGet]
        //[Route("Activity/EditProcessDetails")]
        //public async Task<ActionResult> EditProcess(int id)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.EditProcess(id);
        //    return Ok(obj);
        //}

        //Update Existing Activity
        //[HttpPost]
        //[Route("Activity/UpdateProcessDetails")]
        //public async Task<ActionResult> UpdateProcess(EntityProcess data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.UpdateProcess(data);
        //    return Ok(obj);
        //}

        //Update Existing Activity
        [HttpGet]
        [Route("Activity/DeleteProcessDetails")]
        public async Task<ActionResult> DeleteProcess(int id)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.DeleteProcess(id);
            return Ok(obj);
        }

        //[HttpPost]
        //[Route("Activity/AddUploadedProcessDetails")]
        //public async Task<ActionResult> AddUploadedProcessDetails(proclist data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.AddUploadedProcessDetails(data);
        //    return Ok(obj);
        //}
    }
}
