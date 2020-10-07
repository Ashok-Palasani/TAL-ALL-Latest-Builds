using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblNcProgramTransferMain
    {
        public int NcProgramTransferId { get; set; }
        public int? McId { get; set; }
        public string ProgramNumber { get; set; }
        public int? VersionNumber { get; set; }
        public string ProgramData { get; set; }
        public DateTime? CreatedDate { get; set; }
        public int? CreatedBy { get; set; }
        public bool? IsDeleted { get; set; }
        public bool? FromCnc { get; set; }
        public int? PartId { get; set; }
        public int? OperationNo { get; set; }
        public string CorrectedDate { get; set; }

        public virtual Tblmachinedetails Mc { get; set; }
    }
}
