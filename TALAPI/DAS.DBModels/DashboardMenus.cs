using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class DashboardMenus
    {
        public int Id { get; set; }
        public string MenuName { get; set; }
        public string MenuUrl { get; set; }
        public string ImageUrl { get; set; }
        public int? MenuId { get; set; }
        public string ColourDiv { get; set; }
        public string Style { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
