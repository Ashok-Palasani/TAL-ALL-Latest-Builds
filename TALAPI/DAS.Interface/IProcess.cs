using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.Interface
{
   public interface IProcess
    {
        CommonResponse GetProcess();
        EntityModel CreateProcess(EntityProcess data);
       // CommonResponse EditProcess(int id);
       // CommonResponse UpdateProcess(EntityProcess data);
        CommonResponse DeleteProcess(int id);

        //upload process details excel
        CommonResponse AddUploadedProcessDetails(proclist data);
    }
}
