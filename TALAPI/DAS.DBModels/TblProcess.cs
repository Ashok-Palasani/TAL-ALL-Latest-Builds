using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblProcess
    {
        public int ProcessId { get; set; }
        public string ProcessName { get; set; }
        public string ProcessDescc { get; set; }
        public int Isdeleted { get; set; }
        public DateTime Createdon { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
    }
}
