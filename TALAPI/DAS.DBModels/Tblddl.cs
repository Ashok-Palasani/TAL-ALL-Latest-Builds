using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblddl
    {
        public int Ddlid { get; set; }
        public string WorkCenter { get; set; }
        public string WorkOrder { get; set; }
        public string OperationNo { get; set; }
        public string OperationDesc { get; set; }
        public string MaterialDesc { get; set; }
        public string PartName { get; set; }
        public string TargetQty { get; set; }
        public string Type { get; set; }
        public string Project { get; set; }
        public DateTime? Maddate { get; set; }
        public string MaddateInd { get; set; }
        public string PreOpnEndDate { get; set; }
        public string DaysAgeing { get; set; }
        public string OperationsOnHold { get; set; }
        public string ReasonForHold { get; set; }
        public DateTime? DueDate { get; set; }
        public string FlagRushInd { get; set; }
        public string SplitWo { get; set; }
        public int IsCompleted { get; set; }
        public int IsDeleted { get; set; }
        public int DeliveredQty { get; set; }
        public int ScrapQty { get; set; }
        public DateTime? InsertedOn { get; set; }
        public DateTime? CorrectedDate { get; set; }
        public string F26 { get; set; }
        public string F27 { get; set; }
        public int? IsWoselected { get; set; }
        public int? Bhmiid { get; set; }
        public int PlantCode { get; set; }
        public string PlannedStartDateTime { get; set; }
        public string PlannedFinishDateTime { get; set; }
        public int IsPcs { get; set; }
        public string PcpNo { get; set; }
    }
}
