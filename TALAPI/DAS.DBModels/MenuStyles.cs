using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class MenuStyles
    {
        public int Id { get; set; }
        public string StyleName { get; set; }
        public int? IsDeleted { get; set; }
    }
}
