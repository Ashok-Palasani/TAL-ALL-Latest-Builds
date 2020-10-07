using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblProgramTransferDetailsMaster
    {
        public int PtdMid { get; set; }
        public int? PlantId { get; set; }
        public int? ProgramType { get; set; }
        public string IpAddress { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int? Port { get; set; }
        public string Domain { get; set; }
        public int Isdeleted { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public string MachineProgramPath { get; set; }
        public string MachineInvNo { get; set; }
        public string MachineModel { get; set; }
        public string ControllerType { get; set; }
        public string MachineDispName { get; set; }
        public string MachineMake { get; set; }
        public int? Shopid { get; set; }
        public int? CellId { get; set; }

        public virtual Tblcell Cell { get; set; }
        public virtual Tblplant Plant { get; set; }
        public virtual TblprogramType ProgramTypeNavigation { get; set; }
        public virtual Tblshop Shop { get; set; }
    }
}
