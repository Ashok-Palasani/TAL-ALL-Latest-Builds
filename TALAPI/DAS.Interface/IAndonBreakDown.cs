using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.Interface
{
    public interface IAndonBreakDown
    {
        CommonResponse BreakDownStart(int machineId);
        CommonResponse BreakDownEnd(int machineId);
        CommonResponse BreakDownStartTxtFile(int machineId);
        CommonResponse BreakDownEndTxtFile(int machineId);
    }
}
