using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class EntityReport
    {
        public int plantId { get; set; }
        public int shopId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public string fromDate { get; set; }
        public string toDate { get; set; }
    }

    public class NoCompliance
    {
        public string fromDate { get; set; }
        public string toDate { get; set; }
        public int machineId { get; set; }
    }

    public class OEEDeckFormat
    {
        public string fromDate { get; set; }
        public string toDate { get; set; }
        public int plantId { get; set; }
        public int shopId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
    }

    public class OEEloadSQLTable
    {
        public DateTime weekStartDateTime { get; set; }
        public int machineId { get; set; }
        public string machineInvNo { get; set; }
        public string machineDispName { get; set; }
        public decimal targetOEE { get; set; }
        public int durationInMinForWeek { get; set; }
        public decimal avaliabilityForWeek { get; set; }
        public int weekNumber { get; set; }
        public decimal timeTookMinForWeek { get; set; }
        public decimal PerformanceForWeek { get; set; }
        public decimal OEE { get; set; }

    }

    public class OEELossloadSQLTable
    {
        public int OEECatID { get; set; }
        public string OEECatName { get; set; }
        public int targetInHrs { get; set; }
        public int machineID { get; set; }
        public int weekNumber { get; set; }
        public int weekStartyear { get; set; }
        public int lossDurationinMinForWeek { get; set; }
        public int totalLossDurationinMinForWeek { get; set; }
        public int overrallLossPOBDMLL { get; set; }
        public decimal overallContribution { get; set; }
    }

    public class ShiftList
    {
       public int shfitId { get; set; }
       public string shiftName { get; set; }
       public DateTime shiftStartDateTime { get; set; }
       public DateTime shiftEndDateTime { get; set; }
    }
}
