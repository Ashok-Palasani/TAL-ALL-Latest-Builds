using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class SidebarMenus
    {
        public int Id { get; set; }
        public string SubMenuName { get; set; }
        public string MenuName { get; set; }
        public string SubMenuUrl { get; set; }
        public string MenuUrl { get; set; }
        public string ImageUrl { get; set; }
        public int? MenuId { get; set; }
        public int? DisplayOrder { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public string CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
