using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class SplitDurationEntity
    {
        public class ModeStartEndDateTime
        {
            public int tempModeId { get; set; }
            public int modeId { get; set; }
            public string startDate { get; set; }
            public string endDate { get; set; }
            public string startTime { get; set; }
            public string endTime { get; set; }
            public string message { get; set; }
        }

        public class UpdateModeDetails
        {
            public int modeId { get; set; } 
            public string endTime { get; set; }
        }

        public class AddReasonsDetails
        {
            public int tempModeId { get; set; }
            public int modeId { get; set; }
            public string mode { get; set; }
            public int lossCodeId { get; set; }
            public int machineId { get; set; }
        }

        public class SendMailDetails
        {
            public int plantId { get; set; }
            public int shopId { get; set; }
            public int cellId { get; set; }
            public int machineId { get; set; }
            public string date { get; set; }
            public int reasonId { get; set; }
            public string tempModeIds { get; set; }
            public string checkIds { get; set; }
        }

        public class ApproveReject
        {
            public int machineId { get; set; }
            public string machineName { get; set; }
            public string mode { get; set; }
            public string reason1 { get; set; }
            public string reason2 { get; set; }
            public string reason3 { get; set; }
            public string startTime { get; set; }
            public string endTime { get; set; }
            public int tempModeId { get; set; }
        }

        //Get Index Data
        public class ModeIndex
        {
            public int tempModeId { get; set; }
            public string startDate { get; set; }
            public string endDate { get; set; }
            public string mode { get; set; }
            public string reasonLevel1 { get; set; }
            public string reasonLevel2 { get; set; }
            public string reasonLevel3 { get; set; }
            public string sentApproval { get; set; }
            public string accpetReject { get; set; }
        }

        public class reasonDetails
        {
            public string reasonLevel1 { get; set; }
            public string reasonLevel2 { get; set; }
            public string reasonLevel3 { get; set; }
        }

        public class AddIndividualMode
        {
            public int tempModeId { get; set; }
            public string mode { get; set; }
            public int lossCodeId { get; set; }
        }

        public class ModeStartEndDateTimeWithReasons
        {
            public int tempModeId { get; set; }
            public int modeId { get; set; }
            public string startDate { get; set; }
            public string endDate { get; set; }
            public string startTime { get; set; }
            public string endTime { get; set; }
            public string mode { get; set; }
            public int res1 { get; set; }
            public int res2 { get; set; }
            public int res3 { get; set; }
            public string res1Name { get; set; }
            public string res2Name { get; set; }
            public string res3Name { get; set; }
            public int overAllSaved { get; set; }
        }

        public class IndexDetails
        {
            public int tempModeId { get; set; }
            public string correctedDate { get; set; }
            public string plantName { get; set; }
            public string shopName { get; set; }
            public string cellName { get; set; }
            public string machineName { get; set; }
            public string firstApproval { get; set; }
            public string secondApproval { get; set; }
        }

        public class SplitDetails
        {
            public int tempModeId { get; set; }
            public string tempModeIds { get; set; }
            public string endTime { get; set; }
        }

        public class DeleteTempMode
        {
            public int deleteTempModeId { get; set; }
            public string deleteTempModeIds { get; set; }
        }

        public class UpdateTempMode
        {
            public int updateTempModeId { get; set; }
            public string updateTempModeIds { get; set; }
            public string endTime { get; set; }
        }

        public class LoginInfo
        {
            public string userName { get; set; }
            public string password { get; set; }
            public string tempModeIds { get; set; }
            public int machineId { get; set; }
        }
    }
}
