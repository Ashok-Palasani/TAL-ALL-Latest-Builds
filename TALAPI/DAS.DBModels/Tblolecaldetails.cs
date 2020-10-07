using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblolecaldetails
    {
        public int Oleid { get; set; }
        public int? Opid { get; set; }
        public int? OpWorkingDuration { get; set; }
        public int? LossDuration { get; set; }
        public int? MachineId { get; set; }
        public string CorrectedDate { get; set; }
        public string Ottime { get; set; }
        public int? Shift { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int Isdeleted { get; set; }
        public DateTime? DeletedDate { get; set; }
        public double? Blue { get; set; }
        public double? Green { get; set; }
        public double? SettingTime { get; set; }
        public double? Roalossess { get; set; }
        public double? SummationOfSctvsPp { get; set; }
        public double? ScrapQtyTime { get; set; }
        public double? ReWotime { get; set; }
    }
}
