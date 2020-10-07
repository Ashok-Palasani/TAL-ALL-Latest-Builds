using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    // autosuggest the workorder number
    public class EntityHMIWrongQty
    {
        public string workOrderNo { get; set; }
    }

    // autosuggest operation number
    public class AutoSuggestOperationNo
    {
        public string operationNo { get; set; }
    }

    // for getting data
    public class Getdata
    {
        public string WoNo { get; set; }
        public string OpNo { get; set; }
        public string id { get; set; }
       public string unCheckId { get; set; }
        //public string id { get; set; } // button disable in mail
        //public string Date { get; set; }
    }

    //for Partial Finish
    public class pf
    {
        public string wqtyids { get; set; }
        public int wqtyid { get; set; }
        public int Delqty { get; set; }
    }

    // for getting data and handling the double click mail
    public class GetMaildata
    {
        public string WoNo { get; set; }
        public string OpNo { get; set; }
        public string id { get; set; } // button disable in mail
        //public string Date { get; set; }
    }

    //getting all the id and value in strings
    public class GetIdsValues
    {
        public string wQtyHmiIds { get; set; }
        public string values { get; set; }
    }


    // for getting data with reason for reject
    public class GetdataWithRejectReason
    {
        public string WoNo { get; set; }
        public string OpNo { get; set; }
        public string id { get; set; }
        public string unCheckId { get; set; }
        //public string Date { get; set; }
        public int reasonId { get; set; }
    }

    // for display the data
    public class WODetails
    {
        public string WoNo { get; set; }
        public string OpNo { get; set; }
        public string partno { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string Project { get; set; }
        public string prodfai { get; set; }
        public string WoQty { get; set; }
        public bool jf { get; set; }
        public bool pf { get; set; }
        public string ProcessQty { get; set; }
        public string DeliveredQty { get; set; }
        public string MachineName { get; set; }
        public int WQtyhmiid { get; set; }
    }


    // for display the data in index sheet
    public class WODetailsData
    {
        //public string WoNo { get; set; }
        //public string OpNo { get; set; }
        //public string partno { get; set; }
        //public string StartTime { get; set; }
        //public string EndTime { get; set; }
        //public string Project { get; set; }
        //public string prodfai { get; set; }
        //public string WoQty { get; set; }
        //public string ProcessQty { get; set; }
        //public string DeliveredQty { get; set; }
        public string MachineName { get; set; }
        public int WQtyhmiid { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string CorrectedDate { get; set; }
        public string firstApproval { get; set; }
        public string secondApproval { get; set; }
        public int machineId { get; set; }
    }

    //update qty and validateqty
    public class ValidateQTYUpdate
    {
        public int qty { get; set; }
        public int WQtyhmiid { get; set; }
    }

    //update qty and validateqty
    public class AutoSuggestOpetaion
    {
        public string workOrderNo { get; set; }
        public string operationNo { get; set; }
    }
}
