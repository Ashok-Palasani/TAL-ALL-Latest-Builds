using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
   public class EntityActivity
    {
        public int ActivityId { get; set; }
        public string ActivityName { get; set; }
        public string ActivityDesc { get; set; }
        public int plantId { get; set; }
        public int shopId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public int processId { get; set; }
        public string optionalAct { get; set; }
    }
    public class EntityActivityForGet
    {
        public int ActivityId { get; set; }
        public string ActivityName { get; set; }
        public string ActivityDesc { get; set; }
        public int plantId { get; set; }
        public string plantName { get; set; }
        public int shopId { get; set; }
        public string shopName { get; set; }
        public int cellId { get; set; }
        public string cellName { get; set; }
        public int processId { get; set; }
        public string processName { get; set; }
        public string optionalAct { get; set; }
        public int machineId { get; set; }
        public string machineName { get; set; }
    }

    //public class activityclass
    //{
    //    public string Name { get; set; }
    //    public string Description { get; set; }
    //    public string cellName { get; set; }
    //    public string plantName { get; set; }
    //    public string shopName { get; set; }
    //    public string processName { get; set; }
    //    public string IsOptional { get; set; }
    //}

    //public class actlist
    //{
    //    public List<activityclass> activitylist { get; set; }
    //}
    public class EditActivity
    {
        public string ActivityName { get; set; }
        public string ActivityDesc { get; set; }
    }

    public class GetProcess
    {
        public int processId { get; set; }
        public string processName { get; set; }
    }

    public class validateprocess
    {
        public int cellId { get; set; }
        public int ProcessId { get; set; }
    }

    public class validateprocessCell
    {
        public int cellId { get; set; }
        public int ProcessId { get; set; }
        public string activityName { get; set; }
    }

    public class EditEmployee
    {
        public string EmpName { get; set; }
        public string EmpDesignation { get; set; }
        public int EmpId { get; set; }
        public string EmpRole { get; set; }
    }

    public class GetEmployee
    {
        public int EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpDesignation { get; set; }
        public int EmpRole { get; set; }
        public int PlantId { get; set; }
        public int ShopId { get; set; }
        public int CellId { get; set; }
        public string CantactNo { get; set; }
        public string emailId { get; set; }
    }

    public class getReasonTable
    {
        public string MachineName { get; set; }
        public int ncid { get; set; }
        public string reasonLevel1 { get; set; }
        public string reasonLevel2 { get; set; }
        public string reasonLevel3 { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
    }

    //Added New code

    public class tcflist
    {
        public string type { get; set; }
        public List<tcfclass> approvedlist { get; set; }
    }

    public class HTclass
    {
        public string WONo { get; set; }
        public string OPNo { get; set; }
        public string PartNo { get; set; }
        public string HTNo { get; set; }
    }
    public class HTlist
    {
        public string type { get; set; }
        public List<HTclass> htlist { get; set; }
    }

    public class PcpNoCustom
    {
        public string partNumber { get; set; }
        public string specialProcessInvolved { get; set; }
        public string textToBeMentionedInRCWithPCPNumbers { get; set; }
    }

    public class PcpDetails
    {
        public string type { get; set; }
        public List<PcpNoCustom> PCPNoForHeatCycle { get; set; }
    }

    public class ScrapqntList
    {
        public string order { get; set; }
        public string workCtr { get; set; }
        public string shortDescription { get; set; }
        public string opAc { get; set; }
        public string PostGDate { get; set; }
        public string yield { get; set; }
        public string setup { get; set; }
        public string mach { get; set; }
        public string labour { get; set; }
        public string rev { get; set; }
        public string ccldConf { get; set; }
        public string scrapQuantity { get; set; }

    }

    public class Scrapqnty
    {
        public string type { get; set; }
        public List<ScrapqntList> ScrapqntyList { get; set; }
    }
}
