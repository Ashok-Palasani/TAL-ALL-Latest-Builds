using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblshopForSeimen
    {
        public int SShopId { get; set; }
        public string ShopName { get; set; }
        public string ShopDesc { get; set; }
        public int PlantId { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
