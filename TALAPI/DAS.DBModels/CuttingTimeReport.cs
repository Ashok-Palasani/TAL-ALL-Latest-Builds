using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class CuttingTimeReport
    {
        public int CuttingId { get; set; }
        public double Cuttingtimes { get; set; }
        public TimeSpan Cuttingtimeinserted { get; set; }
        public int CuttingShift { get; set; }
        public string CuttingDate { get; set; }
    }
}
