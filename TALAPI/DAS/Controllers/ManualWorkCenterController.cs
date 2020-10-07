using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.ManualWCEntity;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class ManualWorkCenterController : ControllerBase
    {
        private readonly IManualWorkCenter manualWorkCenter;
        public ManualWorkCenterController(IManualWorkCenter _manualWorkCenter)
        {
            manualWorkCenter = _manualWorkCenter;
        }

        [HttpGet]
        [Route("ManualWorkCenter/GetPlants")]
        public async Task<IActionResult> GetPlants()
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = manualWorkCenter.GetPlants();
            return Ok(response);
        }

        [HttpGet]
        [Route("ManualWorkCenter/GetShops")]
        public async Task<IActionResult> GetShops(int plantId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = manualWorkCenter.GetShops(plantId);
            return Ok(response);
        }

        [HttpGet]
        [Route("ManualWorkCenter/GetCells")]
        public async Task<IActionResult> GetCells(int shopId)
        {
            //calling DepartmentDAL busines layer
            CommonResponse response = manualWorkCenter.GetCells(shopId);
            return Ok(response);
        }

        [HttpPost]
        [Route("ManualWorkCenter/AddStandardMachineDetails")]
        public async Task<IActionResult> AddStandardMachineDetails([FromBody]AddMachineandUserDetails data)
        {
            //calling DepartmentDAL busines layer

            GeneralResponse response = new GeneralResponse();

            response = manualWorkCenter.AddStandardMachineDetails(data);
            //return Ok(response);;
            return Ok(response);
        }

        [HttpGet]
        [Route("ManualWorkCenter/GetManualWorkCenterCount")]
        public async Task<IActionResult> GetManualWorkCenterCount(int no, string MachineInvNo)
        {
            //calling DepartmentDAL busines layer

            CommonResponse response = new CommonResponse();

            response = manualWorkCenter.GetManualWorkCenterCount(no, MachineInvNo);
            //return Ok(response);;
            return Ok(response);
        }

        [HttpPost]
        [Route("ManualWorkCenter/AddManualWorkCenterAndUserDetails")]
        public async Task<IActionResult> AddManualWorkCenterAndUserDetails([FromBody]List<AddMachineandUserDetails> datas)
        {
            //calling DepartmentDAL busines layer

            CommonResponse response = new CommonResponse();

            response = manualWorkCenter.AddManualWorkCenterAndUserDetails(datas);
            //return Ok(response);;
            return Ok(response);
        }
    }
}