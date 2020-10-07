using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSprejectreason
    {
        public int Sprid { get; set; }
        public string RejectName { get; set; }
        public string RejectNameDesc { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string CorrectedDate { get; set; }
        public int? Cellid { get; set; }
        public int? Machineid { get; set; }
        public int IsMaint { get; set; }
        public int IsProd { get; set; }
        public int IsTcf { get; set; }
        public int Sphmiid { get; set; }
    }
}
