using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblOeecat
    {
        public int OeecatId { get; set; }
        public string OeecatName { get; set; }
        public double TargetInHrs { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
