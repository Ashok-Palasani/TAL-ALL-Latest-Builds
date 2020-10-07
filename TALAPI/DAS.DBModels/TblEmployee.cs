using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblEmployee
    {
        public int Eid { get; set; }
        public int? EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpDesignation { get; set; }
        public int? EmpRole { get; set; }
        public int? PlantId { get; set; }
        public int? ShopId { get; set; }
        public int? CellId { get; set; }
        public string CantactNo { get; set; }
        public string EmailId { get; set; }
        public int Isdeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
