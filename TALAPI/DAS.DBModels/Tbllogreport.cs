using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tbllogreport
    {
        public int Logid { get; set; }
        public string LogDescription { get; set; }
        public DateTime? LogDate { get; set; }
        public DateTime? LogCapturedTime { get; set; }
        public int? MachineId { get; set; }
        public int? PartNo { get; set; }
        public int? OperationNo { get; set; }
        public int? Version { get; set; }
        public int? UserId { get; set; }
        public string ProgramNumber { get; set; }
        public int IsDeleted { get; set; }
        public string LogInfo { get; set; }
    }
}
