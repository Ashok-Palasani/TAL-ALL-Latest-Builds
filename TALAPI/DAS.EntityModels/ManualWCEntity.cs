using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class ManualWCEntity
    {
        public class AddMachineandUserDetails
        {
            public int machineId { get; set; }
            public string machineInvNo { get; set; }
            public string ipaddress { get; set; }
            public string controllerType { get; set; }
            public string machineModel { get; set; }
            public string machineMake { get; set; }
            public string modelType { get; set; }
            public string machineDispName { get; set; }
            public int? isPcb { get; set; }
            public int? plantId { get; set; }
            public int? shopId { get; set; }
            public int? cellId { get; set; }
            public int? manualWcid { get; set; }
            public int isSyncEnable { get; set; }
            public string password { get; set; }
            public string userName { get; set; }
        }

        public class ManualWorkCenter
        {
            public string manualWorkCenterName { get; set; }
        }

        public class MachineandUserDetails
        {
            public int machineId { get; set; }
            public string machineInvNo { get; set; }
            public string ipaddress { get; set; }
            public string controllerType { get; set; }
            public string machineModel { get; set; }
            public string machineMake { get; set; }
            public string modelType { get; set; }
            public string machineDispName { get; set; }
            public int? isPcb { get; set; }
            public int? plantId { get; set; }
            public int? shopId { get; set; }
            public int? cellId { get; set; }
            public string plantName { get; set; }
            public string shopName { get; set; }
            public string cellName { get; set; }
            public int isSyncEnable { get; set; }
            public int? manualWcid { get; set; }
            public string password { get; set; }
            public string userName { get; set; }
        }
    }
}
