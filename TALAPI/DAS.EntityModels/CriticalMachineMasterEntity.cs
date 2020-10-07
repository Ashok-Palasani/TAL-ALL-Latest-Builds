using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class CriticalMachineMasterEntity
    {
        public class CriticalMachine
        {
            public string checkedMachineIds { get; set; }
            public string unCheckedMachineIds { get; set; }
        }

        public class ViewCrititcalMachine
        {
            public int crititcalMachineId { get; set; }
            public string machineInvNo { get; set; }
            public string machineDispName { get; set; }
            public string plantName { get; set; }
            public string shopName { get; set; }
            public string cellName { get; set; }
            public string date { get; set; }
        }
    }
}
