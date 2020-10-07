using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CriticalMachineMasterEntity;

namespace DAS.Interface
{
    public interface ICriticalMachineMaster
    {
        CommonResponse1 GetPlants();
        CommonResponse1 GetShops(int plantId);
        CommonResponse1 GetCells(int shopId);
        CommonResponseForCount GetMachines(int cellId);
        CommonResponse1 AddCriticalMachines(CriticalMachine data);
        CommonResponse1 ViewCriticalMachines();
        CommonResponse1 DeleteCrititcalMachine(int id);
    }
}
