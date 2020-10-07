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
    public class EmployeeController : Controller
    {
        IEmployee actobj;
        IPlantShopCellData iPlantShopCellData;

        public EmployeeController(IEmployee act, IPlantShopCellData _iPlantShopCellData)
        {
            actobj = act;
            iPlantShopCellData = _iPlantShopCellData;
        }

        //Get Plant Details
        [HttpGet]
        [Route("Employee/GetPlantDetails")]
        public async Task<ActionResult> GetPlant()
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetPlantDetails();
            return Ok(obj);
        }

        //Get Shop Details
        [HttpGet]
        [Route("Employee/GetShopDetails")]
        public async Task<ActionResult> GetShop(int plantId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetShopDetails(plantId);
            return Ok(obj);
        }

        //Get Cell Details
        [HttpGet]
        [Route("Employee/GetCellDetails")]
        public async Task<ActionResult> GetCell(int shopId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iPlantShopCellData.GetCellDetails(shopId);
            return Ok(obj);
        }

        //Get Role Details
        [HttpGet]
        [Route("Employee/GetRoleDetails")]
        public async Task<ActionResult> GetRole()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetRole();
            return Ok(obj);
        }

        //Get Employee Details
        [HttpGet]
        [Route("Employee/GetEmployeeDetails")]
        public async Task<ActionResult> GetEmployeeDet()
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.GetEmployee();
            return Ok(obj);
        }

        //Create New Employee
        [HttpPost]
        [Route("Employee/CreateOrEditEmployeeDetails")]
        public async Task<ActionResult> CreateEmployeeDet(EntityEmployee data)
        {
            EntityModel obj = new EntityModel();
            obj = actobj.CreateEmployee(data);
            return Ok(obj);
        }

        //Get Individual Activity to Edit
        [HttpGet]
        [Route("Activity/EditProcessDetails")]
        public async Task<ActionResult> EditEmployeeDet(int id)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.EditEmployee(id);
            return Ok(obj);
        }

        //Update Existing Employee
        //[HttpPost]
        //[Route("Employee/UpdateEmployeeDetails")]
        //public async Task<ActionResult> UpdateEmployeeDet(EntityEmployee data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = actobj.UpdateEmployee(data);
        //    return Ok(obj);
        //}

        //Update Existing Employee
        [HttpGet]
        [Route("Employee/DeleteEmployeeDetails")]
        public async Task<ActionResult> DeleteEmployeeDet(int id)
        {
            CommonResponse obj = new CommonResponse();
            obj = actobj.DeleteEmployee(id);
            return Ok(obj);
        }


    }
}
