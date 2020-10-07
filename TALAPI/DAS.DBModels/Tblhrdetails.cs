using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblhrdetails
    {
        public int Hid { get; set; }
        public int? Opid { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string CorrectedDate { get; set; }
        public int? DurationInMin { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int Isdeleted { get; set; }
    }
}
