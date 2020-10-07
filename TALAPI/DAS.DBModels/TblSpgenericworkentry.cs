using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblSpgenericworkentry
    {
        public int SpgwentryId { get; set; }
        public int GwcodeId { get; set; }
        public DateTime? StartDateTime { get; set; }
        public DateTime? EndDateTime { get; set; }
        public string CorrectedDate { get; set; }
        public int MachineId { get; set; }
        public string Shift { get; set; }
        public string GwcodeDesc { get; set; }
        public string Gwcode { get; set; }
        public int DoneWithRow { get; set; }
        public string BatchNo { get; set; }
    }
}
