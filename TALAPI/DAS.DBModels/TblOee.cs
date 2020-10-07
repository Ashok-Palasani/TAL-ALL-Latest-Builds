using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblOee
    {
        public int OeeId { get; set; }
        public int? MachineId { get; set; }
        public decimal? StdOee { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
