using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    // getting the data from login details table the timing
    public class EntityNoLogin
    {
        public int plantId { get; set; }
        public int shopiId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public string fromDate { get; set; }
        public string id { get; set; }
        public string unCheckId { get; set; }
        //public string toDate { get; set; }  // for selecting multiple day
    }

    public class NoSplitDetails
    {
        public int noLoginId { get; set; }
        public string noLoginIds { get; set; }
        public string endTime { get; set; }
    }

    public class SendNoLoginDetails
    {
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string correctedDate { get; set; }
        public int machineId { get; set; }
        public int loginDetailsId { get; set; }
    }

    public class NoLoginStartEndDateTime
    {
        public int noLoginId { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public int loginDetailsId { get; set; }
    }

    // send data to the view to showcase
    public class NoLoginDet
    {
        public int noLoginId { get; set; }
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string machineName { get; set; }
        public string workOrderNo { get; set; }
        public string partNo { get; set; }
        public string oprationNo { get; set; }
        public string workOrderQty { get; set; }
        public string project { get; set; }
        public string processedQty { get; set; }
        public string prodFai { get; set; }
        public string operatorId { get; set; }
        public string deleveredQty { get; set; }
        public string shift { get; set; }
        public int isWocenter { get; set; }
        public bool ddl { get; set; }
        public bool jf { get; set; }
        public bool pf { get; set; }
        public bool rwo { get; set; }
        public bool isSplit { get; set; }
        public int loginDetailsId { get; set; }
        //public bool start { get; set; }
    }

    // send data to the view to showcase In index
    public class NoLoginDetData
    {
        public int noLoginId { get; set; }
        //public string startDateTime { get; set; }
        //public string endDateTime { get; set; }
        //public string machineName { get; set; }
        //public string workOrderNo { get; set; }
        //public string partNo { get; set; }
        //public string oprationNo { get; set; }
        //public string workOrderQty { get; set; }
        //public string project { get; set; }
        //public string processedQty { get; set; }
        //public string prodFai { get; set; }
        //public string operatorId { get; set; }
        //public string deleveredQty { get; set; }
        //public string shift { get; set; }
        //public int isWocenter { get; set; }
        //public bool ddl { get; set; }
        //public bool pf { get; set; }
        //public bool rwo { get; set; }
        //public bool isSplit { get; set; }
        public string machineName { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string CorrectedDate { get; set; }
        public string firstApproval { get; set; }
        public string secondApproval { get; set; }
        public int machineId { get; set; }
    }

    // to set the operatorid
    public class NoLoginOperatorDetails
    {
        public int operatorId { get; set; }
        public int noLoginId { get; set; }
    }
    // to set the Shift
    public class NoLoginSetShift
    {
        public string shift { get; set; }
        public int noLoginId { get; set; }
    }

    //Getting the data for ddl
    public class NoLoginDDLCommonResponse
    {
        public bool isTure { get; set; }
        public dynamic response { get; set; }
        public int count { get; set; }
    }

    // For DDl Lasy Loading
    public class NoLoginDDLList
    {
        public int machineId { get; set; }
        public int takeValue { get; set; }
        public int skipeValue { get; set; }
        public int noLoginId { get; set; }
        //public int count { get; set; }
    }

    // for set reworkorder
    public class NoLoginSetReWork
    {
        public int noLoginId { get; set; }
        public int isChecked { get; set; }
    }

    // for selecting workorder
    public class NoLoginSelectWO
    {
        public string ddlId { get; set; }
        public int noLoginId { get; set; }
    }

    //Sending the DDLids and new table primaryKey
    public class NoLoginSendDDLUnAsignedWoId
    {
        public int noLoginId { get; set; }
        public string ddlIds { get; set; }
    }

    // fro storing form ddl to new table
    public class NoLoginStoreToUnAsigned
    {
        public int noLoginId { get; set; }
        public string workOrderNo { get; set; }
        public string partNo { get; set; }
        public string oprationNo { get; set; }
        public string project { get; set; }
        public string workOrderQty { get; set; }
        public string ddlWorkCenter { get; set; }
        public string prodFai { get; set; }
    }

    // for set Splitworkorder
    public class NoLoginSetSplitWork
    {
        public int noLoginId { get; set; }
        public int isChecked { get; set; }
    }

    //jobfinish
    public class NoLoginJobFinish
    {
        public int noLoginId { get; set; }
        public string deliveryQty { get; set; }
    }

    public class NoLoginStoreHMIDetails
    {
        public int noLoginId { get; set; }
        public string deliveryQty { get; set; }
    }

    public class DeleteSplitDuration
    {
        public int deleteNoLoginId { get; set; }
        public string deleteNoLoginIds { get; set; }
    }

    public class UpdateNoLogin
    {
        public int updateNoLoginId { get; set; }
        public string updateNoLoginIds { get; set; }
        public string endTime { get; set; }
    }
}
