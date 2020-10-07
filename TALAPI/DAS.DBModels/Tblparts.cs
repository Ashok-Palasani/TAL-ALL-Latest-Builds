using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblparts
    {
        public int PartId { get; set; }
        public int PartNo { get; set; }
        public string PartDesc { get; set; }
        public string PartName { get; set; }
        public int IdleCycleTime { get; set; }
        public int UnitDesc { get; set; }
        public int IsDeleted { get; set; }
        public string CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }

        public virtual Tblunit UnitDescNavigation { get; set; }
    }
}
