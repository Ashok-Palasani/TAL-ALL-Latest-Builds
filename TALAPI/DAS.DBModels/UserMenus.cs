using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class UserMenus
    {
        public int Id { get; set; }
        public int? MenuId { get; set; }
        public int? RoleId { get; set; }
        public bool? MenuStatus { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
