//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace i_facility_IdleHandlerWithOptimization.TataModel
{
    using System;
    using System.Collections.Generic;
    
    public partial class main_time_rep
    {
        public int MainTimeRepID { get; set; }
        public double CuttingTime { get; set; }
        public double OperatingTime { get; set; }
        public double PowerOnTime { get; set; }
        public string CorrectedDate { get; set; }
        public int Shift { get; set; }
        public int MachineID { get; set; }
        public Nullable<int> MonthNum { get; set; }
        public Nullable<int> WeekNum { get; set; }
        public Nullable<int> YearNum { get; set; }
        public string MonthName { get; set; }
        public string WeekRange { get; set; }
    }
}
