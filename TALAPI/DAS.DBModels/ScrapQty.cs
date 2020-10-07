using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class ScrapQty
    {
        public int ScrapId { get; set; }
        public string WorkOder { get; set; }
        public string WorkCenter { get; set; }
        public string Decription { get; set; }
        public string OperationNo { get; set; }
        public string PostgDate { get; set; }
        public int? Yield { get; set; }
        public decimal? Setup { get; set; }
        public decimal? Mach { get; set; }
        public int? Labour { get; set; }
        public string Rev { get; set; }
        public int? CcldConf { get; set; }
        public int? NoOfscrapQty { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
