using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tbloperatordetails
    {
        public int Opid { get; set; }
        public string Dept { get; set; }
        public string OperatorName { get; set; }
        public string OperatorId { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CorrectedDate { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int? OperatorDesc { get; set; }
        public string OperatorMailId { get; set; }
        public string ContactNo { get; set; }
        public string EmployeeId { get; set; }

        public virtual Tblroles OperatorDescNavigation { get; set; }
    }
}
