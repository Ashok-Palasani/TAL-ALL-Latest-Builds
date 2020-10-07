using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class LossCodeLevel1
    {
        public int LossCodeID { get; set; }
        public string LossCode { get; set; }
    }
    public class RejectReason
    {
        public int RejectID { get; set; }
        public string RejectName { get; set; }
    }
    public class LossCodeLevel2
    {
        public int LossCodeID { get; set; }
        public string LossCode { get; set; }
    }
    public class LossCodeLevel3
    {
        public int LossCodeID { get; set; }
        public string LossCode { get; set; }
    }

    public class NoCodeStartEndDateTime
    {
        public int noCodeId { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public string message { get; set; }
    }

    public class tcfclass
    {
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string Modulename { get; set; }

        public string firstCCList { get; set; }
        public string firstToList { get; set; }
        public string secondCCList { get; set; }
        public string secondToList { get; set; }
        //public int Lossid { get; set; }
        //public int NCID { get; set; }
        //public int reasonLevel1Id { get; set; }
        //public int reasonLevel2Id { get; set; }
        //public int reasonLevel3Id { get; set; }
        public string CorrectedDate { get; set; }
        //public DateTime EndTime { get; set; }
    }
    public class updateLoss
    {
        public int NCID { get; set; }
        public int reasonLevel1Id { get; set; }
        public int reasonLevel2Id { get; set; }
        public int reasonLevel3Id { get; set; }
        public int NoOfReason { get; set; }
    }
    public class reason
    {
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
    }
    public class reasonlevel
    {
        public int reasonLevel1Id { get; set; }
        public int reasonLevel2Id { get; set; }
        public int reasonLevel3Id { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
    }
}
