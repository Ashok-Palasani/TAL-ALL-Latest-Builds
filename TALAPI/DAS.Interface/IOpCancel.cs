using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.OPCancelEntity;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.Interface
{
   public interface IOpCancel
    {
        CommonResponse GetPrevOpCancelDetails();

        CommonResponse GetPrevOpCancelDetails(LsitOPcancelDet Opcanceldet );
        //Get all the success and fail data and show case
        CommonResponse GetListOFOperationNumberFileData(List<UpLoadExcel> data);
        //select and unselect the operatio number
        CommonResponse AcceptRejectOperationNo(AcceptCancel data);

        //Login details for Approval
        CommonResponse1 LoginDetails(LoginInfo data);
        //get the the Canceled list
        CommonResponse GetCancledList();
        //get the the unCanceled list
        CommonResponse GetUnCancledList();
        //Send for approval
        CommonResponse SendForApproval();
        //accept the selected operation cancellation
        CommonResponse AcceptPrvOp(ApprovalClass data);
        //Get the reject reasons
        CommonResponse GetRejectReason();
        //Accept the reject reasons
        CommonResponse RejectPrvOPCan(RejectClass data);
        //Common method for displaying the data in tables
        CommonResponse CommonDtatForDisplay(ApprovalClass data);
        //Index
        CommonResponse Index();
    }
}
