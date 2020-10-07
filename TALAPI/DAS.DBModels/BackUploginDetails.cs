using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class BackUploginDetails
    {
        public long BloginDetailsId { get; set; }
        public long? MachineId { get; set; }
        public string Shift { get; set; }
        public string CorrectedDate { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public bool? IsCompleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public DateTime? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public long? ModifiedBy { get; set; }
        public bool? IsDeleted { get; set; }
        public bool? IsActive { get; set; }
    }
}
