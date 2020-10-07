using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblCriticalMachine
    {
        public int CriticalMachineId { get; set; }
        public int? MachineId { get; set; }
        public string CorrectedDate { get; set; }
        public DateTime? InsertedOn { get; set; }
        public int? InsertedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int? IsCritical { get; set; }
        public int? IsDeleted { get; set; }
    }
}
