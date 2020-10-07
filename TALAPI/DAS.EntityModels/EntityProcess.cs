using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
   public class EntityProcess
    {
            public int processId { get; set; }
            public string processName { get; set; }
            public string processDesc { get; set; }
    }
    public class EditProcess
    {
        public string ProcessName { get; set; }
        public string ProcessDesc { get; set; }
    }

    public class processclass
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }

    public class proclist
    {
        public string type { get; set; }  //New code added
        public List<processclass> processlist { get; set; }
    }

    //New code added

    public class activityclass
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public int machineId { get; set; }
        public string machineName { get; set; }
        public string cellName { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string processName { get; set; }
        public string IsOptional { get; set; }
    }

    public class actlist
    {
        public string type { get; set; }
        public List<activityclass> activitylist { get; set; }
    }
}
