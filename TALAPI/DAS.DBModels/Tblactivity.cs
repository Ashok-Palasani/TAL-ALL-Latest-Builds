using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblactivity
    {
        public int ActivityId { get; set; }
        public string ActivityName { get; set; }
        public string ActivityDescc { get; set; }
        public int Isdeleted { get; set; }
        public DateTime Createdon { get; set; }
        public int CreatedBy { get; set; }
        public int? PlantId { get; set; }
        public int? ShopId { get; set; }
        public int? CellId { get; set; }
        public int? MachineId { get; set; }
        public string OptionalAct { get; set; }
        public int? ProcessId { get; set; }
    }
}
