//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AutoReportMail
{
    using System;
    using System.Collections.Generic;
    
    public partial class operating_time_report_pdf
    {
        public int OperatingID { get; set; }
        public double Operatingtimes { get; set; }
        public Nullable<System.TimeSpan> Operatingtimeinserted { get; set; }
        public int OperatingShift { get; set; }
        public string OperatingDate { get; set; }
        public Nullable<int> MachineID { get; set; }
        public Nullable<int> MonthNum { get; set; }
        public Nullable<int> WeekNum { get; set; }
        public Nullable<int> YearNum { get; set; }
    }
}
