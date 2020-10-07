using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DAS.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        IReport iReport;
        public ReportController(IReport _iReport)
        {
            iReport = _iReport;
        }

        //Machine Status Register Report
        [HttpPost]
        [Route("Report/MachineStatusRegister")]
        public async Task<ActionResult> MachineStatusReport(EntityReport data)
        {
            CommonResponse obj = new CommonResponse();
            obj = iReport.MachineStatusRegister(data);
            return Ok(obj);
        }


        //Manual Wo Start,Confirmation and split output file generation
        //[HttpPost]
        //[Route("Report/ManualWOConfirmationAndSplitAndStart")]
        //public async Task<ActionResult> ManualWOConfirmationAndSplitAndStart(ReportEntity data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    obj = iReport.ManualWOConfirmationAndSplitAndStart(data);
        //    return Ok(obj);
        //}

        [HttpPost]
        [Route("Report/NoComplianceReport")]
        public async Task<IActionResult> NoComplianceReport([FromBody]OEEDeckFormat data)
        {
            //calling DepartmentDAL busines layer

            CommonResponse1 response = new CommonResponse1();

            response = iReport.NoComplianceReport(data);

            //return Ok(response);;
            return Ok(response);
        }


        [HttpPost]
        [Route("Report/OEEDeckFormatReport")]
        public async Task<IActionResult> OEEDeckFormatReport([FromBody]OEEDeckFormat data)
        {
            //calling DepartmentDAL busines layer

            CommonResponse1 response = new CommonResponse1();

            response = iReport.OEEDeckFormatReport(data);

            //return Ok(response);;
            return Ok(response);
        }

        [HttpPost]
        [Route("Report/OEEDeckFormatLossReasonReport")]
        public async Task<IActionResult> OEEDeckFormatLossReasonReport([FromBody]OEEDeckFormat data)
        {
            //calling DepartmentDAL busines layer

            CommonResponse1 response = new CommonResponse1();

            response = iReport.OEEDeckFormatLossReasonReport(data);

            //return Ok(response);;
            return Ok(response);
        }

        [HttpGet]
        [Route("Report/UpdateLoginDetails")]
        public async Task<IActionResult> UpdateLoginDetails(long machineId)
        {
            //calling DepartmentDAL busines layer

            CommonResponse response = new CommonResponse();

            response = iReport.UpdateLoginDetails(machineId);

            //return Ok(response);;
            return Ok(response);
        }

       
    }
}