﻿using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tbldncroles
    {
        public int RoleId { get; set; }
        public string RoleDesc { get; set; }
        public string RoleType { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
