using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class Smtpdetails
    {
        public int SmtpId { get; set; }
        public int Port { get; set; }
        public string Host { get; set; }
        public bool EnableSsl { get; set; }
        public bool UseDefaultCredentials { get; set; }
        public bool Certificate { get; set; }
        public bool IsDeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string FromMailId { get; set; }
        public string Password { get; set; }
        public string EmailId { get; set; }
        public string ConnectType { get; set; }
        public int? TcfModuleId { get; set; }
    }
}
