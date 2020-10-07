using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Menus
    {
        public int Id { get; set; }
        public string MenuName { get; set; }
        public int? IsDashboard { get; set; }
        public int? IsSideMenubar { get; set; }
        public int? DisplayOrder { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
