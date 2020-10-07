using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
   public class EntityEmployee
    {
        public int EId { get; set; }
        public int EmpId { get; set; }
        public string CantactNo { get; set; }
        public string EmpName { get; set; }
        public string EmpDesignation { get; set; }
        public int EmpRole { get; set; }
        public int PlantId { get; set; }
        public int ShopId { get; set; }
        public int CellId { get; set; }
        public string emailId { get; set; }
    }

    public class GetEmp
    {
        public int EId { get; set; }
        public int EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpDesignation { get; set; }
        public int EmpRole { get; set; }
        public string PlantName { get; set; }
        public string rollName { get; set; }
        public int PlantId { get; set; }
        public int ShopId { get; set; }
        public int CellId { get; set; }
        public string ShopName { get; set; }
        public string CellName { get; set; }
        public string CantactNo { get; set; }
        public string emailId { get; set; }
    }
}
