using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblDdlstatus
    {
        public int DdlstatusId { get; set; }
        public string StatusMessage { get; set; }
        public int Status { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public int IsDeleted { get; set; }
    }
}
