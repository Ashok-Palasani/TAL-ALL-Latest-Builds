using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
  public  class CommonEntity
    {
        public class HoldIdleCodeResponse
        {
            public bool isStatus { get; set; }
            public dynamic response { get; set; }
            public string MachineDisplayName { get; set; }
            public int Level { get; set; }
            public int LossId { get; set; }
            public string losscode { get; set; }
            public string StartTime { get; set; }
            public string ItsLastLevel { get; set; }
           public string errorMsg { get; set; }
        }

        public class MachineForOEECal
        {
            public int MachineID { get; set; }
        }
        


        public class Responce
        {
            public string message { get; set; }
        }

        public class AddWOCommonResponse
        {
            public bool isStatus { get; set; }
            public dynamic response { get; set; }
            public List<Responce> responce { get; set; }
        }

        public class GeneralResponse
        {
            public bool isStatus { get; set; }
            public dynamic response { get; set; }
            public string color { get; set; }
            public int machineId { get; set; }
            public string machineInvNo { get; set; }
        }

        public class BreakDownCodeResponse
        {
            public bool isStatus { get; set; }
            public dynamic response { get; set; }
            public string MachineDisplayName { get; set; }
            public int Level { get; set; }
            public int BreakdownId { get; set; }
            public string ItsLastLevel { get; set; }
            public string BreakdownCode { get; set; }
            public string BreakdownStartTime { get; set; }
            public string errorMsg { get; set; }
        }

        public class CommonResponseForList
        {
            public bool isStatus { get; set; }
            public dynamic response { get; set; }
            public string MachineDisplayName { get; set; }
        }
    }
}
