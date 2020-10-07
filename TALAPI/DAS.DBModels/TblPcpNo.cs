using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblPcpNo
    {
        public int PcpId { get; set; }
        public string PartNo { get; set; }
        public string SpecialProcessInvolved { get; set; }
        public string PcpNo { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
