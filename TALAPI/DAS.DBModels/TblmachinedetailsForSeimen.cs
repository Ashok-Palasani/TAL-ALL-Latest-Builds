using System;
using System.Collections.Generic;

namespace DAS.DBModels
{
    public partial class TblmachinedetailsForSeimen
    {
        public int SmachineId { get; set; }
        public string MachineInvNo { get; set; }
        public string Ipaddress { get; set; }
        public string MachineType { get; set; }
        public string ControllerType { get; set; }
        public string InsertedOn { get; set; }
        public int InsertedBy { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public int? ModifiedBy { get; set; }
        public int? IsDeleted { get; set; }
        public string MachineModel { get; set; }
        public string MachineMake { get; set; }
        public string ModelType { get; set; }
        public string MachineDispName { get; set; }
        public int IsParameters { get; set; }
        public string ShopNo { get; set; }
        public int? IsPcb { get; set; }
        public int? IsLevel { get; set; }
        public int? PlantId { get; set; }
        public int? ShopId { get; set; }
        public int? CellId { get; set; }
        public string DeletedDate { get; set; }
        public int IsNormalWc { get; set; }
        public int? ManualWcid { get; set; }
        public int IsSyncEnable { get; set; }
        public int? IsCriticalMachine { get; set; }
        public string EquipmentNum { get; set; }
        public int? IsNestedSheetMachine { get; set; }
        public int IsDlversion { get; set; }
    }
}
