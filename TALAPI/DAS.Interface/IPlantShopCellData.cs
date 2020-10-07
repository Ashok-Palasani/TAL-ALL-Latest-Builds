using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonResponseWithMachineName;

namespace DAS.Interface
{
    public interface IPlantShopCellData
    {
        CommonResponse GetPlantDetails();
        CommonResponse GetShopDetails(int plantId);
        CommonResponse GetCellDetails(int shopId);
        CommonResponse GetMachineDetails(int cellId);
        CommonResponse GetPlantShopCellMachineNames(PlantShopCellMachineId data);
        CommonResponse GetAllDetailsBasedOnPlantShopCell(PlantShopCellGet data);
        CommonResponse GetAllDetailsBasedOnPlantShopCellMachine(PlantShopCellGet data);
        CommonResponse GetAllDetailsBasedOnPlantShopCellProcess(PlantShopCellGet data);
    }
}
