//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace MimicsUpdation
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblTempLiveLossOfEntry
    {
        public int tempLossId { get; set; }
        public Nullable<int> messageCodeId { get; set; }
        public Nullable<System.DateTime> startDateTime { get; set; }
        public Nullable<System.DateTime> endDateTime { get; set; }
        public Nullable<System.DateTime> entryTime { get; set; }
        public string correctedDate { get; set; }
        public int machineID { get; set; }
        public string shift { get; set; }
        public string messageDesc { get; set; }
        public string messageCode { get; set; }
        public int isUpdate { get; set; }
        public int doneWithRow { get; set; }
        public Nullable<int> isStart { get; set; }
        public Nullable<int> isScreen { get; set; }
        public int forRefresh { get; set; }
        public Nullable<int> tempModeId { get; set; }
    }
}
