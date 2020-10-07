using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblFtpDetails
    {
        public int Ptid { get; set; }
        public int PlantId { get; set; }
        public int ShopId { get; set; }
        public int CellId { get; set; }
        public int MachineId { get; set; }
        public int ProgType { get; set; }
        public string IpAddress { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int Port { get; set; }
        public string Domain { get; set; }
        public DateTime CreatedOn { get; set; }
        public int CreatedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int Isdeleted { get; set; }
    }
}
