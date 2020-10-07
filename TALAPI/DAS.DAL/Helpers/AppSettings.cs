using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.DAL.Helpers
{
    public class AppSettings
    {
        public string Secret { get; set; }
        public string CommonEmail { get; set; }
        public string DocumentEmail { get; set; }
        public string ResetLinkURL { get; set; }
        public string ImageUrl { get; set; }
        public string ImageUrlSave { get; set; }
        public string CandidateUploadURL { get; set; }
        public string DefaultConnection { get; set; }
    }
}
