using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblcellForSiemen
    {
        public int SCellId { get; set; }
        public string CellName { get; set; }
        public string CellDesc { get; set; }
        public int? PlantId { get; set; }
        public int ShopId { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
