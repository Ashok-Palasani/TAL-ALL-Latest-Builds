using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;



namespace DAS.Interface
{
   public interface IPreactorSchedule
    {
        CommonResponse GetPreactorSchedule();

        CommonResponse CreatePreactorSchedule(PreactorEntity data);
      
        CommonResponse DeletePreactorSchedule(int id);

        // New code added
        CommonResponse1 AddUploadedProcessDetails(proclist data);
        CommonResponse1 AddUploadedActivityDetails(actlist data);
        CommonResponse1 AddUploadedApprovedMasterDetails(tcflist data);
        CommonResponse1 AddUploadedHeattreamentDetails(HTlist data);
        CommonResponse1 UploadPcpNo(PcpDetails data);
        CommonResponse1 GetPcpNo();
        CommonResponse1 UploadScrapQuantity(Scrapqnty data);
        CommonResponse1 GetScrapQuantity();
    }
}
