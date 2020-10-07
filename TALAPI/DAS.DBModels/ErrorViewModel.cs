using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.DBModels
{
    public class ErrorViewModel
    {
        public string RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}
