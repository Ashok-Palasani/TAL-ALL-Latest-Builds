using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSpGeneric
    {
        public int Id { get; set; }
        public string OperatorId { get; set; }
        public string BatchNumber { get; set; }
        public int? GenericCode { get; set; }
        public int? MachineId { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
    }
}
