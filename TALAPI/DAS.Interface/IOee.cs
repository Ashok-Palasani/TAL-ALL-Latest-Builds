using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.OeeEntity;

namespace DAS.Interface
{
    public interface IOee
    {
        CommonResponse1 AddAndUpdateOee(OeeDetails data);
        CommonResponse1 DeleteOee(int oeeId);
        CommonResponse1 ViewOeeDetails();
        CommonResponse1 ViewOeeDetailsById(int oeeId);
        CommonResponse1 UpdateOee(int oeeId, decimal stdOee);
    }
}
