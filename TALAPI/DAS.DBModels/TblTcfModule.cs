using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblTcfModule
    {
        public int TcfModuleId { get; set; }
        public string TcfModuleName { get; set; }
        public string TcfModuleDesc { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int? IsDeleted { get; set; }
    }
}
