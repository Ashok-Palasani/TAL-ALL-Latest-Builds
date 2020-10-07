using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.OeeEntity;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class OeeController : ControllerBase
    {
        private readonly IOee oee;
        public OeeController(IOee _oee)
        {
            oee = _oee;
        }

        /// <summary>
        /// Add And Update Oee
        /// </summary>
        /// <param name="machineId"></param>
        /// <param name="stdOee"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("Oee/AddAndUpdateOee")]
        public async Task<IActionResult> AddAndUpdateOee(OeeDetails data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = oee.AddAndUpdateOee(data);
            return Ok(response);
        }

        /// <summary>
        /// Update Oee
        /// </summary>
        /// <param name="oeeId"></param>
        /// <param name="stdOee"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("Oee/UpdateOee")]
        public async Task<IActionResult> UpdateOee(int oeeId, decimal stdOee)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = oee.UpdateOee(oeeId, stdOee);
            return Ok(response);
        }

        /// <summary>
        /// Delete Oee
        /// </summary>
        /// <param name="oeeId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("Oee/DeleteOee")]
        public async Task<IActionResult> DeleteOee(int oeeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = oee.DeleteOee(oeeId);
            return Ok(response);
        }

        /// <summary>
        /// View Oee Details
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Oee/ViewOeeDetails")]
        public async Task<IActionResult> ViewOeeDetails()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = oee.ViewOeeDetails();
            return Ok(response);
        }

        /// <summary>
        /// View Oee Details By Id
        /// </summary>
        /// <param name="oeeId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("Oee/ViewOeeDetailsById")]
        public async Task<IActionResult> ViewOeeDetailsById(int oeeId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = oee.ViewOeeDetailsById(oeeId);
            return Ok(response);
        }
    }
}