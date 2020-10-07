using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblPrevOperationCancel
    {
        public int OpcancelId { get; set; }
        public string ProductionOrder { get; set; }
        public string Operation { get; set; }
        public int IsCancelled { get; set; }
        public int? Qty { get; set; }
        public string CorrectedDate { get; set; }
        public string WorkCenter { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public string PartNumber { get; set; }
    }
}
