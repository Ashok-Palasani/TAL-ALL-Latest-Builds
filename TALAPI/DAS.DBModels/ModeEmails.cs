using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class ModeEmails
    {
        public int Eid { get; set; }
        public string Tomailids { get; set; }
        public string Ccmailids { get; set; }
        public int Isdeleted { get; set; }
        public DateTime? Createdon { get; set; }
        public int Createdby { get; set; }
        public string MessageMode { get; set; }
    }
}
