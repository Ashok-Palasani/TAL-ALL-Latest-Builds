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
    public class AndonBreakdownController : ControllerBase
    {
        IAndonBreakDown iAndonBreakDown;

        public AndonBreakdownController(IAndonBreakDown _iAndonBreakDown)
        {
            iAndonBreakDown = _iAndonBreakDown;
        }

        //On Click of breakdownstart
        [HttpGet]
        [Route("AndonBreakdown/BreakDownStart")]
        public async Task<ActionResult>BreakDownStart(int machineId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iAndonBreakDown.BreakDownStart(machineId);
            return Ok(obj);
        }

        //On Click of breakdownEnd
        [HttpGet]
        [Route("AndonBreakdown/BreakDownEnd")]
        public async Task<ActionResult> BreakDownEnd(int machineId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iAndonBreakDown.BreakDownEnd(machineId);
            return Ok(obj);
        }


        //On Click of breakdownstart txt file
        [HttpGet]
        [Route("AndonBreakdown/BreakDownStartTxt")]
        public async Task<ActionResult> BreakDownStartTxt(int machineId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iAndonBreakDown.BreakDownStartTxtFile(machineId);
            return Ok(obj);
        }

        //On Click of breakdownEnd txt file
        [HttpGet]
        [Route("AndonBreakdown/BreakDownEndTxt")]
        public async Task<ActionResult> BreakDownEndTxt(int machineId)
        {
            CommonResponse obj = new CommonResponse();
            obj = iAndonBreakDown.BreakDownEndTxtFile(machineId);
            return Ok(obj);
        }
    }
}