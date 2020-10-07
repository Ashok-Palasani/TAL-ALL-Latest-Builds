using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblTcfApprovedMaster
    {
        public int TcfApprovedMasterId { get; set; }
        public int? TcfModuleId { get; set; }
        public string FirstApproverToList { get; set; }
        public string FirstApproverCcList { get; set; }
        public string SecondApproverToList { get; set; }
        public string SecondApproverCcList { get; set; }
        public int? PlantId { get; set; }
        public int? ShopId { get; set; }
        public int? CellId { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
