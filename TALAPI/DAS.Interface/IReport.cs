using DAS.EntityModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.Interface
{
    public interface IReport
    {
        // for generating the machine status register report
        CommonResponse MachineStatusRegister(EntityReport data);

        // for generating the Manual WO Start,Confirmation,Split report
       // CommonResponse ManualWOConfirmationAndSplitAndStart(ReportEntity data);

        // NoComplianceReport
        CommonResponse1 NoComplianceReport(OEEDeckFormat data);

        CommonResponse1 OEEDeckFormatReport(OEEDeckFormat data);

        CommonResponse1 OEEDeckFormatLossReasonReport(OEEDeckFormat data);

        CommonResponse UpdateLoginDetails(long machineId);
    }
}
