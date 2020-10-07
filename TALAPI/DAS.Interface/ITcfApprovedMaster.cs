using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.TcfApprovedMasterEntity;

namespace DAS.Interface
{
    public interface ITcfApprovedMaster
    {
        CommonResponse1 AddAndEditTcfApprovedMaster(AddAndEditTcfMaster data);
        CommonResponse1 ViewMultipleTcfApprovedMaster();
        CommonResponse1 ViewMultipleTcfApprovedMasterById(int tcfMasterId);
        CommonResponse1 DeleteMultipleTcfApprovedMaster(int tcfMasterId);
        CommonResponse1 GetModules();

        //Upload excel
        //CommonResponse AddUploadedApprovedMasterDetails(tcflist data);
    }
}
