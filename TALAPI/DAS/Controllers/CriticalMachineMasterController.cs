using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CriticalMachineMasterEntity;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class CriticalMachineMasterController : ControllerBase
    {
        private readonly ICriticalMachineMaster criticalMachineMaster;
        public CriticalMachineMasterController(ICriticalMachineMaster _criticalMachineMaster)
        {
            criticalMachineMaster = _criticalMachineMaster;
        }

        /// <summary>
        /// Get Plants
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/GetPlants")]
        public async Task<IActionResult> GetPlants()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.GetPlants();
            return Ok(response);
        }

        /// <summary>
        /// Get Shops
        /// </summary>
        /// <param name="plantId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/GetShops")]
        public async Task<IActionResult> GetShops(int plantId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.GetShops(plantId);
            return Ok(response);
        }

        /// <summary>
        /// Get Cells
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/GetCells")]
        public async Task<IActionResult> GetCells(int shopId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.GetCells(shopId);
            return Ok(response);
        }

        /// <summary>
        /// Get Machines
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/GetMachines")]
        public async Task<IActionResult> GetMachines(int cellID)
        {
            //calling DepartmentDAL busines layer
            CommonResponseForCount response = criticalMachineMaster.GetMachines(cellID);
            return Ok(response);
        }

        /// <summary>
        /// Add Critical Machines
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("CriticalMachine/AddCriticalMachines")]
        public async Task<IActionResult> AddCriticalMachines(CriticalMachine data)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.AddCriticalMachines(data);
            return Ok(response);
        }

        /// <summary>
        /// View Critical Machines
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/ViewCriticalMachines")]
        public async Task<IActionResult> ViewCriticalMachines()
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.ViewCriticalMachines();
            return Ok(response);
        }

        /// <summary>
        /// Delete Crititcal Machine
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CriticalMachine/DeleteCrititcalMachine")]
        public async Task<IActionResult> DeleteCrititcalMachine(int id)
        {
            //calling DepartmentDAL busines layer
            CommonResponse1 response = criticalMachineMaster.DeleteCrititcalMachine(id);
            return Ok(response);
        }
    }
}