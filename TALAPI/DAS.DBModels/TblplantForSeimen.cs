using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblplantForSeimen
    {
        public int SPlantId { get; set; }
        public string PlantName { get; set; }
        public string PlantDesc { get; set; }
        public int IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
