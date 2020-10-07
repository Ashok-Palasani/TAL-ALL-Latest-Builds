using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class DayStEndTime
    {
        public int DayId { get; set; }
        public TimeSpan DayStart { get; set; }
        public TimeSpan DayEnd { get; set; }
    }
}
