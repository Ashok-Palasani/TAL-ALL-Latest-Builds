using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class ModeLog
    {
        public int Mid { get; set; }
        public string MailSentDateTime { get; set; }
        public int MachineId { get; set; }
        public int? CreatedBy { get; set; }
        public DateTime? CreatedOn { get; set; }
    }
}
