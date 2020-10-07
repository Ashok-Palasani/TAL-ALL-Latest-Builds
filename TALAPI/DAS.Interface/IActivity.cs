using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.Interface
{
   public interface IActivity
    {
        CommonResponse GetActivity();
        //Get Process List
        CommonResponse GetProcessList();
        EntityModel CreateActivity(EntityActivity data);
        //CommonResponse EditActivityDet(int id);
        CommonResponse UpdateActivity(EntityActivity data);
        CommonResponse DeleteActivity(int id);
        //Validating the Activity whether its already exists for that perticular cell and process
        CommonResponse validateActivity(validateprocessCell data);
        //Validating the Process whether its already exists for that perticular cell
        CommonResponse validateprocess(int cellId, int processId);

        //To Upload Activity excel
        //CommonResponse AddUploadedActivityDetails(actlist data);
    }
}
