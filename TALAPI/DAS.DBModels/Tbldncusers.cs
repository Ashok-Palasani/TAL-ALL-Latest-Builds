using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tbldncusers
    {
        public int UserId { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int PrimaryRole { get; set; }
        public int SecondaryRole { get; set; }
        public string DisplayName { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int? MachineId { get; set; }
    }
}
