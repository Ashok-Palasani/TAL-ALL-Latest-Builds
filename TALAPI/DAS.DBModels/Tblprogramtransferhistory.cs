using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Tblprogramtransferhistory
    {
        public int Pthid { get; set; }
        public int? MachineId { get; set; }
        public int? UserId { get; set; }
        public string ProgramName { get; set; }
        public DateTime? UploadedTime { get; set; }
        public DateTime? ReturnTime { get; set; }
        public int? ReturnStatus { get; set; }
        public string ReturnDesc { get; set; }
        public int? IsDeleted { get; set; }
        public int IsCompleted { get; set; }
        public int? Version { get; set; }
        public int? PartId { get; set; }
        public int? OperationNo { get; set; }
        public string Correcteddate { get; set; }
        public string DeletedDate { get; set; }
        public string Message { get; set; }
    }
}
