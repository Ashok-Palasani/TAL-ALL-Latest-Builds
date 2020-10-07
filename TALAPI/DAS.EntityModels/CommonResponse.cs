using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class CommonResponse
    {
        public bool isTure { get; set; }
        public dynamic response { get; set; }
        public dynamic errorMsg { get; set; }
    }

    public class CommonResponse1
    {
        public bool isStatus { get; set; }
        public dynamic response { get; set; }
        public bool isTure { get; set; } //New code added
    }

    public class GeneralResponse1
    {
        public bool isStatus { get; set; }
        public dynamic response { get; set; }
        public string tempModeIds { get; set; }
    }

    public class CommonResponseForCount
    {
        public bool isStatus { get; set; }
        public dynamic response { get; set; }
        public int count { get; set; }
    }

    public class CommonResponsewithEror
    {
        public bool isTure { get; set; }
        public dynamic response { get; set; }
        public dynamic errorMsg { get; set; }
        //public bool isMultipleIdle { get; set; }
        //public string batchNo { get; set; }
    }

    public class CommonResponseWithMachinedesscName
    {
        public bool isTure { get; set; }
        public bool IsIdle { get; set; }
        public bool IsBreakdown { get; set; }
        public bool IsPM { get; set; }
        public bool IsHold { get; set; }
        public bool isGeneric { get; set; }
        public bool Isrework { get; set; }
        public dynamic response { get; set; }
        public string MacDispName { get; set; }
        public string Shift { get; set; }
        public int processName { get; set; }
        public int activityName { get; set; }
        public string Opid { get; set; }
        public bool set { get; set; }
    }


    public class CommonResponseWithMachineName
    {
        public bool isTure { get; set; }
        public bool IsIdle { get; set; }
        public bool IsBreakdown { get; set; }
        public bool IsPM { get; set; }
        public bool IsHold { get; set; }
        public bool isGeneric { get; set; }
        public bool Isrework { get; set; }
        public dynamic response { get; set; }
        public string MacDispName
        {
            get; set;
        }

        public class BatchDetWithMachineName
        {
            public int MachineId { get; set; }
            public string BatchNo { get; set; }
        }

        public class PlantEntity
        {
            public int plantId { get; set; }
            public string plantName { get; set; }
        }

        public class ShopEntity
        {
            public int shopId { get; set; }
            public string shopName { get; set; }
        }

        public class CellEntity
        {
            public int cellId { get; set; }
            public string cellName { get; set; }
        }

        public class MachineEntity
        {
            public int machineId { get; set; }
            public string machineName { get; set; }
        }

        public class ProcessEntity
        {
            public int processId { get; set; }
            public string processName { get; set; }
        }

        public class PlantShopCellMachineId
        {
            public int plantId { get; set; }
            public int shopId { get; set; }
            public int cellId { get; set; }
            public int machineId { get; set; }
        }

        public class PlantShopCellGet
        {
            public int plantId { get; set; }
            public int shopId { get; set; }
            public int cellId { get; set; }
            public int machineId { get; set; }
        }

        public class PlantShopCellList
        {
            public dynamic plantList { get; set; }
            public dynamic shopList { get; set; }
            public dynamic cellList { get; set; }

        }

        public class PlantShopCellMachineList
        {
            public dynamic plantList { get; set; }
            public dynamic shopList { get; set; }
            public dynamic cellList { get; set; }
            public dynamic machineList { get; set; }
            public dynamic processList { get; set; }
        }

        public class PlantShopCellMachineIdWithName
        {
            public int plantId { get; set; }
            public string plantName { get; set; }
            public int shopId { get; set; }
            public string shopName { get; set; }
            public int cellId { get; set; }
            public string cellName { get; set; }
            public int machineId { get; set; }
            public string machineName { get; set; }
        }
    }
}
