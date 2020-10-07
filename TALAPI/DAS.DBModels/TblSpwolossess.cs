using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSpwolossess
    {
        public int SpWolossesId { get; set; }
        public int? Hmiid { get; set; }
        public int? LossId { get; set; }
        public string BatchNo { get; set; }
        public string ActivityName { get; set; }
        public string ProcessName { get; set; }
        public string LossName { get; set; }
        public decimal? LossDuration { get; set; }
        public int? Level { get; set; }
        public int? LossCodeLevel1Id { get; set; }
        public string LossCodeLevel1Name { get; set; }
        public int? LossCodeLevel2Id { get; set; }
        public string LossCodeLevel2Name { get; set; }
        public DateTime? InsertedOn { get; set; }
        public int IsDeleted { get; set; }
    }
}
