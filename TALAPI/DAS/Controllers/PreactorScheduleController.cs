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
    public class PreactorScheduleController : Controller
    {
        IPreactorSchedule actobj;

        public PreactorScheduleController(IPreactorSchedule act)
        {
            actobj = act;
        }

        //Get PreactorSchedule Details
        [HttpGet]
        [Route("PreactorSchedule/GetPreactorScheduleDetails")]
        public async Task<ActionResult> GetPreactorSchedule()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetPreactorSchedule();
            return Ok(obj);
        }

        //Create New PreactorSchedule
        [HttpPost]
        [Route("PreactorSchedule/CreatePreactorSchedule")]
        public async Task<ActionResult> CreatePreactorSchedule(PreactorEntity data)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.CreatePreactorSchedule(data);
            return Ok(obj);
        }


        //Update Existing PreactorSchedule
        [HttpGet]
        [Route("PreactorSchedule/DeletePreactorScheduleDetails")]
        public async Task<ActionResult> DeletePreactorSchedule(int id)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.DeletePreactorSchedule(id);
            return Ok(obj);
        }

        //New Code added

        /// <summary>
        /// Add Uploaded Process Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/AddUploadedProcessDetails")]
        public async Task<ActionResult> AddUploadedProcessDetails(proclist data)
#pragma warning restore CS1998 // This async method lacks 'await' operators and will run synchronously. Consider using the 'await' operator to await non-blocking API calls, or 'await Task.Run(...)' to do CPU-bound work on a background thread.
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.AddUploadedProcessDetails(data);
            return Ok(obj);
        }

        /// <summary>
        /// Add Uploaded Activity Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/AddUploadedActivityDetails")]
        public async Task<ActionResult> AddUploadedActivityDetails(actlist data)
#pragma warning restore CS1998 // This async method lacks 'await' operators and will run synchronously. Consider using the 'await' operator to await non-blocking API calls, or 'await Task.Run(...)' to do CPU-bound work on a background thread.
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.AddUploadedActivityDetails(data);
            return Ok(obj);
        }

        /// <summary>
        /// Add Uploaded Approved Master Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/AddUploadedApprovedMasterDetails")]
        public async Task<ActionResult> AddUploadedApprovedMasterDetails(tcflist data)
#pragma warning restore CS1998 // This async method lacks 'await' operators and will run synchronously. Consider using the 'await' operator to await non-blocking API calls, or 'await Task.Run(...)' to do CPU-bound work on a background thread.
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.AddUploadedApprovedMasterDetails(data);
            return Ok(obj);
        }

        /// <summary>
        /// Add Uploaded Heat treament Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/AddUploadedHeattreamentDetails")]
        public async Task<ActionResult> AddUploadedHeattreamentDetails(HTlist data)
#pragma warning restore CS1998 // This async method lacks 'await' operators and will run synchronously. Consider using the 'await' operator to await non-blocking API calls, or 'await Task.Run(...)' to do CPU-bound work on a background thread.
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.AddUploadedHeattreamentDetails(data);
            return Ok(obj);
        }

        /// <summary>
        /// Upload Pcp No
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/UploadPcpNo")]
        public async Task<ActionResult> UploadPcpNo(PcpDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.UploadPcpNo(data);
            return Ok(obj);
        }

        /// <summary>
        /// Get Pcp No
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("PreactorSchedule/GetPcpNo")]
        public async Task<ActionResult> GetPcpNo()
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.GetPcpNo();
            return Ok(obj);
        }

        /// <summary>
        /// Upload Scrap Quantity
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("PreactorSchedule/UploadScrapQuantity")]
        public async Task<ActionResult> UploadScrapQuantity(Scrapqnty data)
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.UploadScrapQuantity(data);
            return Ok(obj);
        }

        /// <summary>
        /// Get Scrap Quantity
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("PreactorSchedule/GetScrapQuantity")]
        public async Task<ActionResult> GetScrapQuantity()
        {
            CommonResponse1 obj = new CommonResponse1();
            obj = actobj.GetScrapQuantity();
            return Ok(obj);
        }
    }
}

