using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class ScrapQtyDetails
    {
        public int Id { get; set; }
        public string WorkCenter { get; set; }
        public string WoNo { get; set; }
        public string OperationNo { get; set; }
        public string PostgDate { get; set; }
        public int? Yield { get; set; }
        public int? ScrapQty { get; set; }
        public decimal? StdHrs { get; set; }
        public decimal? StandardOperatingHrs { get; set; }
        public decimal? TotalRejectionHrs { get; set; }
        public int? IsDeleted { get; set; }
        public DateTime? CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public decimal? QualityFactor { get; set; }
    }
}
