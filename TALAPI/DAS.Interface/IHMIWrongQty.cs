using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
    public interface IHMIWrongQty
    {
        #region Auto Suggest

        // Get Workorder Details
        CommonResponse GetWorkOrderDetails(string workOrderNo);

        // Get operation Details
        CommonResponse GetOperationDetails(AutoSuggestOpetaion data);

        #endregion

        //Index method
        CommonResponse IndexWOQty();

        // get the job finish and partial finsh data
        CommonResponse GetWOJFPFDetails(Getdata data);

        //Partial Finish
        CommonResponsewithEror PartialFinish(pf data);

        //Job Finish
        CommonResponsewithEror JobFinish(pf data);

        //Login details for Approval
        CommonResponse1 LoginDetails(LoginInfo data);

        // validate the quantity and update previous with jf to pf updation logic
        CommonResponse ValidateQtyUpdate(ValidateQTYUpdate data);

        //validate the quantity and with the selected workorder list
        CommonResponse ValidateQtyData(GetIdsValues data);

        //Send for approval
        CommonResponse SendApproval(Getdata data);

        //get all the data for showcase
        //CommonResponse GetAllData(Getdata data);
        CommonResponse GetAllData(GetMaildata data);

        //accept the data for updated workorder
        CommonResponse AcceptWoQtyData(Getdata data);

        //Get Reject Reasons
        CommonResponse GetRejectReasons();

        //reject the data for updated workorder
        CommonResponse RejectWoQtyData(GetdataWithRejectReason data);
    }
}
