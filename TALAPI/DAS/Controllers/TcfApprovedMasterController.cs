using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.TcfApprovedMasterEntity;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class TcfApprovedMasterController : ControllerBase
    {
        private readonly ITcfApprovedMaster tcfApprovedMaster;
        public TcfApprovedMasterController(ITcfApprovedMaster _tcfApprovedMaster)
        {
            tcfApprovedMaster = _tcfApprovedMaster;
        }

        /// <summary>
        /// Add And Edit Tcf Approved Master
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("TcfMaster/AddAndEditTcfApprovedMaster")]
        public async Task<IActionResult> AddAndEditTcfApprovedMaster(AddAndEditTcfMaster data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcfApprovedMaster.AddAndEditTcfApprovedMaster(data);
            return Ok(response);
        }

        /// <summary>
        /// View Multiple Tcf Approved Master
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("TcfMaster/ViewMultipleTcfApprovedMaster")]
        public async Task<IActionResult> ViewMultipleTcfApprovedMaster()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcfApprovedMaster.ViewMultipleTcfApprovedMaster();
            return Ok(response);
        }

        /// <summary>
        /// View Multiple Tcf Approved Master By Id
        /// </summary>
        /// <param name="tcfMasterId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("TcfMaster/ViewMultipleTcfApprovedMasterById")]
        public async Task<IActionResult> ViewMultipleTcfApprovedMasterById(int tcfMasterId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcfApprovedMaster.ViewMultipleTcfApprovedMasterById(tcfMasterId);
            return Ok(response);
        }

        /// <summary>
        /// Delete Multiple Tcf Approved Master
        /// </summary>
        /// <param name="tcfMasterId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("TcfMaster/DeleteMultipleTcfApprovedMaster")]
        public async Task<IActionResult> DeleteMultipleTcfApprovedMaster(int tcfMasterId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcfApprovedMaster.DeleteMultipleTcfApprovedMaster(tcfMasterId);
            return Ok(response);
        }

        [HttpGet]
        [Route("TcfMaster/GetModules")]
        public async Task<IActionResult> GetModules()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = tcfApprovedMaster.GetModules();
            return Ok(response);
        }

        //[HttpPost]
        //[Route("TcfMaster/AddUploadedApprovedMasterDetails")]
        //public async Task<ActionResult> AddUploadedApprovedMasterDetails(tcflist data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = tcfApprovedMaster.AddUploadedApprovedMasterDetails(data);
        //    return Ok(obj);
        //}
    }
}