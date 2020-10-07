using DAS.DBModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
   public class BatchProcessingEntity
    {
        public class BatchDetails
        {
            public int IsWorkOrder { get; set; }
            public string DdlIds { get; set; }
            public string UniqueBatchNo { get; set; }
            public int MachineId { get; set; }
        }

        public class AddHoldCodes
        {
            public string Hmiid { get; set; }
            public int HoldCodeId { get; set; }
            public int MachineId { get; set; }
        }

        public class AddIdleCodes
        {
            public int LossCodeID { get; set; }
            public int MachineId { get; set; }
            public bool isStart { get; set; }
        }

        public class HoldList
        {
            public List<Tblmanuallossofentry> HoldListDetailsWO { get; set; }
            public List<Tbllivelossofentry> HoldListDetailsWC { get; set; }
        }

        //Anajli Code

        //public class HoldCodeDetails
        //{
        //    public int HoldCodeID { get; set; }
        //    public int MachineID { get; set; }
        //    public int Level { get; set; }
        //}

       public class IdleCodeDetails
        {
            public int LossCodeID { get; set; }
            public int MachineID { get; set; }
            public int Level { get; set; }
            public bool IsTrue { get; set; }
        }

        public class AddBreakDownCodes
        {
            public int LossCodeID { get; set; }
            public int MachineId { get; set; }
            public int Level { get; set; }
            public bool IsTrue { get; set; }
        }

       
    }
}
