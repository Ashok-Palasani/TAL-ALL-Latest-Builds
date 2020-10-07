using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class TcfApprovedMasterEntity
    {
        public class AddAndEditTcfMaster
        {
            public int TcfApprovedMasterId { get; set; }
            public int TcfModuleId { get; set; }
            public string FirstApproverToList { get; set; }
            public string FirstApproverCcList { get; set; }
            public string SecondApproverToList { get; set; }
            public string SecondApproverCcList { get; set; }
            public int PlantId { get; set; }
            public int ShopId { get; set; }
            public int CellId { get; set; }
        }

        //public class tcfclass
        //{
        //    public string firstToList { get; set; }
        //    public string firstCCList { get; set; }
        //    public string secondToList { get; set; }
        //    public string secondCCList { get; set; }
        //    public string cellName { get; set; }
        //    public string plantName { get; set; }
        //    public string shopName { get; set; }
        //    public string Modulename { get; set; }
        //}

        //public class tcflist
        //{
        //    public List<tcfclass> approvedlist { get; set; }
        //}
        public class ViewTcfMaster
        {
            public int TcfApprovedMasterId { get; set; }
            public int? TcfModuleId { get; set; }
            public string FirstApproverToList { get; set; }
            public string[] FirstApproverCcList { get; set; }
            public string SecondApproverToList { get; set; }
            public string[] SecondApproverCcList { get; set; }
            public int PlantId { get; set; }
            public int ShopId { get; set; }
            public int CellId { get; set; }
            public string PlantName { get; set; }
            public string ShopName { get; set; }
            public string CellName { get; set; }
            public string ModuleName { get; set; }
        }

        public class FirstAppCcLists
        {
            public string FirstApproverCcList { get; set; }
        }

        public class SecondAppCcLists
        {
            public string SecondApproverCcList { get; set; }
        }
    }
}
