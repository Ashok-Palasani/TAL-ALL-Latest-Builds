using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblprogramType
    {
        public TblprogramType()
        {
            TblProgramTransferDetailsMaster = new HashSet<TblProgramTransferDetailsMaster>();
        }

        public int Ptypeid { get; set; }
        public string TypeName { get; set; }
        public int Isdeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }

        public virtual ICollection<TblProgramTransferDetailsMaster> TblProgramTransferDetailsMaster { get; set; }
    }
}
