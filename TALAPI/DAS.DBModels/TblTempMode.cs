using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblTempMode
    {
        public int TempModeId { get; set; }
        public int ModeId { get; set; }
        public int MachineId { get; set; }
        public string Mode { get; set; }
        public DateTime InsertedOn { get; set; }
        public int InsertedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string CorrectedDate { get; set; }
        public int IsDeleted { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public string ColorCode { get; set; }
        public int IsCompleted { get; set; }
        public int? DurationInSec { get; set; }
        public int? ModeMonth { get; set; }
        public int? ModeYear { get; set; }
        public int? ModeWeekNumber { get; set; }
        public int? ModeQuarter { get; set; }
        public int? IsSaved { get; set; }
        public int? IsApproved { get; set; }
        public int? RejectReasonId { get; set; }
        public int? IsSendToApproveOrRej { get; set; }
        public DateTime? ApprovedDate { get; set; }
        public DateTime? MailSendDate { get; set; }
        public int? UserId { get; set; }
        public int? ApproveLevel { get; set; }
        public int? OverAllSaved { get; set; }
        public int? RejectLevel { get; set; }
        public int? IsUpdated { get; set; }
        public int? IsApproved1 { get; set; }
        public int? IsUpdateFinal { get; set; }
        public int? IsHold { get; set; }
    }
}
