using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.ManualWCEntity;

namespace DAS.Interface
{
    public interface IManualWorkCenter
    {
        CommonResponse GetPlants();
        CommonResponse GetShops(int plantId);
        CommonResponse GetCells(int shopId);
        GeneralResponse AddStandardMachineDetails(AddMachineandUserDetails data);
        CommonResponse GetManualWorkCenterCount(int no, string MachineInvNo);
        CommonResponse AddManualWorkCenterAndUserDetails(List<AddMachineandUserDetails> datas);
    }
}
