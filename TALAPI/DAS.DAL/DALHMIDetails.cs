using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.DAL
{
    public class DALHMIDetails : IHMIDetails
    {
        i_facility_talContext db = new i_facility_talContext();
        //public string connectionString = "Server=SRKSDEV001-PC\\SQLSERVER17;Database=i_facility_tsal;user id=sa;password=srks4$;";
        public static IConfiguration configuration;

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALHMIDetails));

        public DALHMIDetails(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }

        //Index data

        #region Prv Index        
        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    string sendApp = "", acceptReject = "";
        //    string correctedDate = "";
        //    try
        //    {
        //        List<UnsignedWOData> listUnsignedWOData = new List<UnsignedWOData>();

        //        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1).OrderBy(m => m.Correcteddate).ToList();
        //        if (unWO.Count > 0)
        //        {
        //            foreach (var row in unWO)
        //            {

        //                int machine = Convert.ToInt32(row.Machineid);
        //                bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
        //                if (row.WorkOrderNo != null)
        //                {
        //                    ddl = true;
        //                }
        //                //if (row.Isworkinprogress == 1)
        //                //{
        //                //    jf = true;
        //                //}
        //                if (row.Isworkinprogress == 0)
        //                {
        //                    pf = true;
        //                }
        //                if (row.ReWork == 1)
        //                {
        //                    rwo = true;
        //                }
        //                //if (row.IsStart == 1)
        //                //{
        //                //    start = true;
        //                //}
        //                if (row.IsSplit == 1)
        //                {
        //                    split = true;
        //                }
        //                if (row.SendApprove == 1)
        //                {
        //                    sendApp = "Sent For Approval";
        //                }
        //                if (row.Acceptreject == 1)
        //                {
        //                    acceptReject = "Yes";
        //                }
        //                else if (row.Acceptreject == 2)
        //                {
        //                    acceptReject = "No";
        //                }

        //                if (correctedDate != row.Correcteddate)
        //                {
        //                    correctedDate = row.Correcteddate;
        //                    int machineId = Convert.ToInt32(row.Machineid);
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    UnsignedWOData objUnsignedWOData = new UnsignedWOData();
        //                    objUnsignedWOData.uaWOId = row.Uawoid;
        //                    objUnsignedWOData.machineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.sendApprove = sendApp;
        //                    objUnsignedWOData.accpetReject = acceptReject;
        //                    listUnsignedWOData.Add(objUnsignedWOData);
        //                }
        //                //var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
        //                //UnsignedWOData objUnsignedWOData = new UnsignedWOData();
        //                //objUnsignedWOData.uaWOId = row.Uawoid;
        //                //objUnsignedWOData.startDateTime = row.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
        //                //objUnsignedWOData.endDateTime = row.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
        //                //objUnsignedWOData.machineName = machineDetais.MachineInvNo;
        //                //objUnsignedWOData.workOrderNo = row.WorkOrderNo;
        //                //objUnsignedWOData.partNo = row.PartNo;
        //                //objUnsignedWOData.oprationNo = row.OprationNo;
        //                //objUnsignedWOData.workOrderQty = row.WorkOrderQty;
        //                //objUnsignedWOData.processedQty = row.ProcessedQty;
        //                //objUnsignedWOData.project = row.Project;
        //                //objUnsignedWOData.prodFai = row.ProdFai;
        //                //objUnsignedWOData.operatorId = row.Operatorid;
        //                //objUnsignedWOData.shift = row.Shiftid;
        //                //objUnsignedWOData.deleveredQty = row.DeliveryQty;
        //                //objUnsignedWOData.isWocenter = machineDetais.IsNormalWc;
        //                //objUnsignedWOData.ddl = ddl;
        //                ////objUnsignedWO.start = start;
        //                //objUnsignedWOData.rwo = rwo;
        //                ////objUnsignedWO.jf = jf;
        //                //objUnsignedWOData.pf = pf;
        //                //objUnsignedWOData.isSplit = split;
        //                //objUnsignedWOData.sendApprove = sendApp;
        //                //objUnsignedWOData.accpetReject = acceptReject;
        //                //listUnsignedWOData.Add(objUnsignedWOData);
        //            }
        //            obj.isTure = true;
        //            obj.response = listUnsignedWOData;
        //            //obj.response = listUnsignedWOData.OrderBy(x => x.startDateTime);
        //        }
        //        else
        //        {
        //            obj.isTure = true;
        //            obj.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}


        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    string firstApproval = "", secondApproval = "";
        //    string correctedDate = "";
        //    int machineid = 0;
        //    try
        //    {
        //        List<UnsignedWOData> listUnsignedWOData = new List<UnsignedWOData>();

        //        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1).OrderByDescending(m => m.Correcteddate).ToList();
        //        if (unWO.Count > 0)
        //        {
        //            foreach (var row in unWO)
        //            {

        //                int machine = Convert.ToInt32(row.Machineid);
        //                bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
        //                if (row.WorkOrderNo != null)
        //                {
        //                    ddl = true;
        //                }
        //                //if (row.Isworkinprogress == 1)
        //                //{
        //                //    jf = true;
        //                //}
        //                if (row.Isworkinprogress == 0)
        //                {
        //                    pf = true;
        //                }
        //                if (row.ReWork == 1)
        //                {
        //                    rwo = true;
        //                }
        //                //if (row.IsStart == 1)
        //                //{
        //                //    start = true;
        //                //}
        //                if (row.IsSplit == 1)
        //                {
        //                    split = true;
        //                }
        //                if (row.AcceptReject1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.AcceptReject1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.Acceptreject == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.Acceptreject == 2)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Rejected";
        //                }

        //                if (machineid != row.Machineid)
        //                {
        //                    correctedDate = row.Correcteddate;
        //                    int machineId = Convert.ToInt32(row.Machineid);
        //                    machineid = machineId;
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    UnsignedWOData objUnsignedWOData = new UnsignedWOData();
        //                    objUnsignedWOData.uaWOId = row.Uawoid;
        //                    objUnsignedWOData.machineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.firstApproval = firstApproval;
        //                    objUnsignedWOData.secondApproval = secondApproval;
        //                    listUnsignedWOData.Add(objUnsignedWOData);
        //                }
        //            }
        //            obj.isTure = true;
        //            obj.response = listUnsignedWOData;
        //            //obj.response = listUnsignedWOData.OrderBy(x => x.startDateTime);
        //        }
        //        else
        //        {
        //            obj.isTure = true;
        //            obj.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}


        // get all the work order based on the date and cell,shop,plant,machine

        #endregion

        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();           
        //    string correctedDate = "";
        //    try
        //    {
        //        List<UnsignedWOData> listUnsignedWOData = new List<UnsignedWOData>();

        //        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1).OrderBy(m => m.Correcteddate).ToList();
        //        if (unWO.Count > 0)
        //        {
        //            foreach (var row in unWO)
        //            {
        //                string firstApproval = "", secondApproval = "";
        //                int machine = Convert.ToInt32(row.Machineid);
        //                bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
        //                if (row.WorkOrderNo != null)
        //                {
        //                    ddl = true;
        //                }
        //                //if (row.Isworkinprogress == 1)
        //                //{
        //                //    jf = true;
        //                //}
        //                if (row.Isworkinprogress == 0)
        //                {
        //                    pf = true;
        //                }
        //                if (row.ReWork == 1)
        //                {
        //                    rwo = true;
        //                }
        //                //if (row.IsStart == 1)
        //                //{
        //                //    start = true;
        //                //}
        //                if (row.IsSplit == 1)
        //                {
        //                    split = true;
        //                }
        //                if (row.AcceptReject1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.AcceptReject1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.Acceptreject == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.Acceptreject == 2)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Rejected";
        //                }

        //                if (correctedDate != row.Correcteddate)
        //                {
        //                    correctedDate = row.Correcteddate;
        //                    int machineId = Convert.ToInt32(row.Machineid);
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    UnsignedWOData objUnsignedWOData = new UnsignedWOData();
        //                    objUnsignedWOData.uaWOId = row.Uawoid;
        //                    objUnsignedWOData.machineId = (int)row.Machineid;
        //                    objUnsignedWOData.machineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.firstApproval = firstApproval;
        //                    objUnsignedWOData.secondApproval = secondApproval;
        //                    listUnsignedWOData.Add(objUnsignedWOData);
        //                }
        //                else
        //                {
        //                    var listdata = listUnsignedWOData.Where(m => m.machineId == row.Machineid && m.CorrectedDate == row.Correcteddate).FirstOrDefault();
        //                    if (listdata != null)
        //                    {

        //                    }
        //                    else
        //                    {
        //                        correctedDate = row.Correcteddate;
        //                        int machineId = Convert.ToInt32(row.Machineid);
        //                        var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                        string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                        string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                        string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                        UnsignedWOData objUnsignedWOData = new UnsignedWOData();
        //                        objUnsignedWOData.uaWOId = row.Uawoid;
        //                        objUnsignedWOData.machineId = (int)row.Machineid;
        //                        objUnsignedWOData.machineName = machineName.MachineInvNo;
        //                        objUnsignedWOData.plantName = plantName;
        //                        objUnsignedWOData.shopName = shopName;
        //                        objUnsignedWOData.cellName = cellName;
        //                        objUnsignedWOData.CorrectedDate = correctedDate;
        //                        objUnsignedWOData.firstApproval = firstApproval;
        //                        objUnsignedWOData.secondApproval = secondApproval;
        //                        listUnsignedWOData.Add(objUnsignedWOData);
        //                    }
        //                }
        //            }
        //            obj.isTure = true;
        //            obj.response = listUnsignedWOData;
        //            //obj.response = listUnsignedWOData.OrderBy(x => x.startDateTime);
        //        }
        //        else
        //        {
        //            obj.isTure = true;
        //            obj.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        public CommonResponse Index()
        {
            CommonResponse obj = new CommonResponse();
            string correctedDate = "";
            try
            {
                //GetProcessedQty("1300005445", "149W6920-13P", "20");

                List<UnsignedWOData> listUnsignedWOData = new List<UnsignedWOData>();

                //var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1).OrderBy(m => m.Correcteddate).ToList();
                var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1).OrderByDescending(m => m.Uawoid).ToList();
                if (unWO.Count > 0)
                {
                    foreach (var row in unWO)
                    {
                        string firstApproval = "", secondApproval = "";
                        int machine = Convert.ToInt32(row.Machineid);
                        bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
                        if (row.WorkOrderNo != null)
                        {
                            ddl = true;
                        }
                        //if (row.Isworkinprogress == 1)
                        //{
                        //    jf = true;
                        //}
                        if (row.Isworkinprogress == 0)
                        {
                            pf = true;
                        }
                        if (row.ReWork == 1)
                        {
                            rwo = true;
                        }
                        //if (row.IsStart == 1)
                        //{
                        //    start = true;
                        //}
                        if (row.IsSplit == 1)
                        {
                            split = true;
                        }
                        if (row.AcceptReject1 == 1)
                        {
                            secondApproval = "Aprroved";

                        }
                        else if (row.AcceptReject1 == 2)
                        {
                            secondApproval = "Rejected";

                        }
                        if (row.Acceptreject == 1)
                        {
                            firstApproval = "Aprroved";
                        }
                        else if (row.Acceptreject == 2)
                        {
                            firstApproval = "Rejected";
                        }

                        if (firstApproval == "")
                        {
                            firstApproval = "Mail Sent and Approval is Pending";
                            secondApproval = "";
                        }
                        else if (secondApproval == "")
                        {
                            secondApproval = "Mail Sent and Approval is Pending";
                        }


                        //if (correctedDate != row.Correcteddate)
                        //{
                        //    correctedDate = row.Correcteddate;
                        //    int machineId = Convert.ToInt32(row.Machineid);
                        //    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
                        //    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
                        //    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
                        //    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
                        //    UnsignedWOData objUnsignedWOData = new UnsignedWOData();
                        //    objUnsignedWOData.uaWOId = row.Uawoid;
                        //    objUnsignedWOData.machineId = (int)row.Machineid;
                        //    objUnsignedWOData.machineName = machineName.MachineInvNo;
                        //    objUnsignedWOData.plantName = plantName;
                        //    objUnsignedWOData.shopName = shopName;
                        //    objUnsignedWOData.cellName = cellName;
                        //    objUnsignedWOData.CorrectedDate = correctedDate;
                        //    objUnsignedWOData.firstApproval = firstApproval;
                        //    objUnsignedWOData.secondApproval = secondApproval;
                        //    listUnsignedWOData.Add(objUnsignedWOData);
                        //}
                        //else
                        //{
                        var listdata = listUnsignedWOData.Where(m => m.machineId == row.Machineid && m.CorrectedDate == row.Correcteddate).FirstOrDefault();
                        if (listdata != null)
                        {

                        }
                        else
                        {
                            correctedDate = row.Correcteddate;
                            int machineId = Convert.ToInt32(row.Machineid);
                            var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
                            string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
                            string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
                            string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
                            UnsignedWOData objUnsignedWOData = new UnsignedWOData();
                            objUnsignedWOData.uaWOId = row.Uawoid;
                            objUnsignedWOData.machineId = (int)row.Machineid;
                            objUnsignedWOData.machineName = machineName.MachineInvNo;
                            objUnsignedWOData.plantName = plantName;
                            objUnsignedWOData.shopName = shopName;
                            objUnsignedWOData.cellName = cellName;
                            objUnsignedWOData.CorrectedDate = correctedDate;
                            objUnsignedWOData.firstApproval = firstApproval;
                            objUnsignedWOData.secondApproval = secondApproval;
                            listUnsignedWOData.Add(objUnsignedWOData);
                        }
                        //}
                    }
                    obj.isTure = true;
                    obj.response = listUnsignedWOData;
                    //obj.response = listUnsignedWOData.OrderBy(x => x.startDateTime);
                }
                else
                {
                    obj.isTure = true;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponse GetMachineDetailsHMI(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            List<UnsignedWO> listUnsignedWO = new List<UnsignedWO>();
            try
            {
                string correctedDate = data.fromDate;
                bool result = StoreIntoUnsignedWO(data);
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                if (result)
                {
                    foreach (var machineRow in machineData)
                    {
                        int machine = machineRow.MachineId;
                        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && x.SendApprove == 0 && x.Correcteddate == correctedDate && x.Machineid == machine).ToList();
                        if (unWO.Count != 0)
                        {
                            foreach (var row in unWO)
                            {
                                bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
                                if (row.WorkOrderNo != null)
                                {
                                    ddl = true;
                                }
                                if (row.Isworkinprogress == 1)
                                {
                                    jf = true;
                                }
                                if (row.Isworkinprogress == 0)
                                {
                                    pf = true;
                                }
                                if (row.Isworkorder == 1)
                                {
                                    rwo = true;
                                }
                                //if (row.IsStart == 1)
                                //{
                                //    start = true;
                                //}
                                if (row.IsSplit == 1)
                                {
                                    split = true;
                                }

                                var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                UnsignedWO objUnsignedWO = new UnsignedWO();
                                objUnsignedWO.uaWOId = row.Uawoid;
                                objUnsignedWO.startDateTime = row.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.endDateTime = row.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.machineName = machineDetais.MachineInvNo;
                                objUnsignedWO.workOrderNo = row.WorkOrderNo;
                                objUnsignedWO.partNo = row.PartNo;
                                objUnsignedWO.oprationNo = row.OprationNo;
                                objUnsignedWO.workOrderQty = row.WorkOrderQty;
                                objUnsignedWO.processedQty = row.ProcessedQty;
                                objUnsignedWO.project = row.Project;
                                objUnsignedWO.prodFai = row.ProdFai;
                                objUnsignedWO.operatorId = row.Operatorid;
                                objUnsignedWO.shift = row.Shiftid;
                                objUnsignedWO.deleveredQty = row.DeliveryQty;
                                objUnsignedWO.isWocenter = machineDetais.IsNormalWc;
                                objUnsignedWO.ddl = ddl;
                                //objUnsignedWO.start = start;
                                objUnsignedWO.rwo = rwo;
                                objUnsignedWO.jf = jf;
                                objUnsignedWO.pf = pf;
                                objUnsignedWO.isSplit = split;
                                listUnsignedWO.Add(objUnsignedWO);
                            }
                            obj.isTure = true;
                            obj.response = listUnsignedWO.OrderBy(x => x.startDateTime);
                        }
                    }
                }
                else
                {
                    foreach (var machineRow in machineData)
                    {
                        int machine = machineRow.MachineId;
                        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && x.SendApprove == 0 && x.Correcteddate == correctedDate && x.Machineid == machine).ToList();
                        if (unWO.Count > 0)
                        {
                            foreach (var row in unWO)
                            {
                                bool ddl = false, jf = false, pf = false, rwo = false, start = false, split = false;
                                if (row.WorkOrderNo != null)
                                {
                                    ddl = true;
                                }
                                if (row.Isworkinprogress == 1)
                                {
                                    jf = true;
                                }
                                if (row.Isworkinprogress == 0)
                                {
                                    pf = true;
                                }
                                if (row.Isworkorder == 1)
                                {
                                    rwo = true;
                                }
                                if (row.IsStart == 1)
                                {
                                    start = true;
                                }
                                if (row.IsSplit == 1)
                                {
                                    split = true;
                                }

                                var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                UnsignedWO objUnsignedWO = new UnsignedWO();
                                objUnsignedWO.uaWOId = row.Uawoid;
                                objUnsignedWO.startDateTime = row.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.endDateTime = row.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.machineName = machineDetais.MachineInvNo;
                                objUnsignedWO.workOrderNo = row.WorkOrderNo;
                                objUnsignedWO.partNo = row.PartNo;
                                objUnsignedWO.oprationNo = row.OprationNo;
                                objUnsignedWO.workOrderQty = row.WorkOrderQty;
                                objUnsignedWO.processedQty = row.ProcessedQty;
                                objUnsignedWO.project = row.Project;
                                objUnsignedWO.prodFai = row.ProdFai;
                                objUnsignedWO.operatorId = row.Operatorid;
                                objUnsignedWO.shift = row.Shiftid;
                                objUnsignedWO.deleveredQty = row.DeliveryQty;
                                objUnsignedWO.isWocenter = machineDetais.IsNormalWc;
                                objUnsignedWO.ddl = ddl;
                                //objUnsignedWO.start = start;
                                objUnsignedWO.rwo = rwo;
                                objUnsignedWO.jf = jf;
                                objUnsignedWO.pf = pf;
                                objUnsignedWO.isSplit = split;
                                listUnsignedWO.Add(objUnsignedWO);
                            }
                            obj.isTure = true;
                            obj.response = listUnsignedWO.OrderBy(x => x.startDateTime);
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.NoItemsFound;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponse DeletePreviousNoCode(EntityHMIDetails data)
        {
            CommonResponse entity = new CommonResponse();
            try
            {

                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                List<Lossdet> obj1 = new List<Lossdet>();
                foreach (var machineRow in machineData)
                {
                    int updateLevel = 1;
                    var getMailIdsLevel = new List<TblTcfApprovedMaster>();

                    getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 2 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
                    if (getMailIdsLevel.Count == 0)
                    {
                        getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 2 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                    }
                    foreach (var rowMail in getMailIdsLevel)
                    {
                        if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                        {
                            updateLevel = 2;
                        }
                    }

                    var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Correcteddate == data.fromDate && x.Machineid == data.machineId).ToList();
                    if (unWO.Count > 0)
                    {
                        foreach (var lossrow in unWO)
                        {
                            if (lossrow.Acceptreject == 1)
                            {
                                entity.isTure = false;
                                entity.errorMsg = "Mail Sent, and First level is Approved";
                                break;
                            }
                            else if (lossrow.SendApprove == 1)
                            {
                                entity.isTure = false;
                                entity.errorMsg = "Mail Sent, and Approval is pending";
                                break;
                            }
                        }
                    }
                    else if (unWO.Count == 0)
                    {
                        bool check = false;
                        foreach (var row in unWO)
                        {
                            db.Tblunasignedwo.Remove(row);
                            db.SaveChanges();
                            check = true;
                        }
                        if (check)
                        {
                            StoreIntoUnsignedWO(data);
                        }
                    }

                }
                if (obj1.Count != 0)
                {
                    entity.isTure = true;
                    entity.response = obj1;
                }
                else
                {
                    entity.isTure = false;
                    entity.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                entity.isTure = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        //Start the WO
        public CommonResponse StartWO(int uaWoID)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var uaWoData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoID).FirstOrDefault();
                if (uaWoData != null)
                {
                    uaWoData.IsStart = 1;
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.FailureMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //get all the timing for the day in tblhimiscreen 
        //public bool StoreIntoUnsignedWO(EntityHMIDetails data)
        //{
        //    bool ret = false;
        //    try
        //    {
        //        DataTable fillData = new DataTable();
        //        string dayTimingStartTime = Convert.ToString(db.Tbldaytiming.Where(x => x.IsDeleted == 0).Select(x => x.StartTime).FirstOrDefault());
        //        int cellId = data.cellId;
        //        int machId = data.machineId;
        //        int shopId = data.shopiId;
        //        int plantId = data.plantId;
        //        string startDate = data.fromDate;
        //        //string endDate = data.toDate; // for selecting multiple day
        //        string endDate = data.fromDate;
        //        double dayDifference = Convert.ToDateTime(endDate).Subtract(Convert.ToDateTime(startDate)).TotalDays;
        //        var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        if (data.machineId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.shopiId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }

        //        //var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        foreach (var machRow in machineData)
        //        {
        //            int machieID = machRow.MachineId;
        //            for (int i = 0; i <= dayDifference; i++)
        //            {
        //                DateTime temp = Convert.ToDateTime(startDate).AddDays(i);
        //                string correctedDate = temp.ToString("yyyy-MM-dd");
        //                var unWODet = db.Tblunasignedwo.Where(x => x.Isupdate == 0 && x.Machineid == machieID && x.Correcteddate == correctedDate).ToList();
        //                if (unWODet.Count == 0)
        //                {
        //                    string dayStart = temp.ToString("yyyy-MM-dd");
        //                    //var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == machieID && x.CorrectedDate == dayStart && x.Status == 2 && x.IsWorkInProgress == 1).Select(x => new { x.Date, x.Time, x.CorrectedDate }).ToList();
        //                    var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == machieID && x.CorrectedDate == dayStart && (x.Status == 2 || x.Status == 1) && (x.IsWorkInProgress == 1 || x.IsWorkInProgress == 0)).Select(x => new { x.Date, x.Time, x.CorrectedDate }).ToList();
        //                    for (int j = 0; j <= hmiData.Count(); j++)
        //                    {
        //                        SendHMIDetails objSendHMIDetails = new SendHMIDetails();
        //                        if (j == 0) // start time row
        //                        {
        //                            dayStart = dayStart + " " + dayTimingStartTime;
        //                            string dayStartTbl = Convert.ToString(hmiData[j].Date);
        //                            double getMinutes = Convert.ToDateTime(dayStartTbl).Subtract(Convert.ToDateTime(dayStart)).TotalMinutes;
        //                            if (dayStart == dayStartTbl)
        //                            {
        //                                // nothing to do
        //                            }
        //                            else if (getMinutes >= 1)
        //                            {
        //                                objSendHMIDetails.startDateTime = dayStart;
        //                                objSendHMIDetails.endDateTime = dayStartTbl;
        //                                objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
        //                                objSendHMIDetails.machineId = machieID;
        //                            }
        //                        }
        //                        else if (j == hmiData.Count())// ending time row
        //                        {
        //                            string previousEndTime = Convert.ToString(hmiData[j - 1].Time);
        //                            string tempStartTime = Convert.ToDateTime(startDate).AddDays(1).ToString("yyyy-MM-dd") + " " + dayTimingStartTime;
        //                            string presentStartTime = Convert.ToDateTime(tempStartTime).AddSeconds(-1).ToString("yyyy-MM-dd HH:mm:ss");
        //                            double getMinutes = Convert.ToDateTime(presentStartTime).Subtract(Convert.ToDateTime(previousEndTime)).TotalMinutes;
        //                            if (presentStartTime == previousEndTime)
        //                            {
        //                                // nothing to do
        //                            }
        //                            else if (getMinutes >= 1)
        //                            {
        //                                objSendHMIDetails.startDateTime = previousEndTime;
        //                                objSendHMIDetails.endDateTime = presentStartTime;
        //                                objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
        //                                objSendHMIDetails.machineId = machieID;
        //                            }
        //                        }
        //                        else if (j > 0)
        //                        {
        //                            string previousEndTime = Convert.ToString(hmiData[j - 1].Time);
        //                            string presentStartTime = Convert.ToString(hmiData[j].Date);
        //                            double getMinutes = Convert.ToDateTime(presentStartTime).Subtract(Convert.ToDateTime(previousEndTime)).TotalMinutes;
        //                            if (presentStartTime == previousEndTime)
        //                            {
        //                                // nothing to do
        //                            }
        //                            else if (getMinutes >= 1)
        //                            {
        //                                objSendHMIDetails.startDateTime = previousEndTime;
        //                                objSendHMIDetails.endDateTime = presentStartTime;
        //                                objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
        //                                objSendHMIDetails.machineId = machieID;
        //                            }
        //                        }
        //                        if (objSendHMIDetails.machineId != 0)
        //                        {
        //                            ret = InsertToUnAsignedWO(objSendHMIDetails);
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return ret;
        //}

        //inserting into new table with start time and endtime

        //get all the timing for the day in tblhimiscreen 

        public bool StoreIntoUnsignedWO(EntityHMIDetails data)
        {
            bool ret = false;
            try
            {
                DataTable fillData = new DataTable();
                string dayTimingStartTime = Convert.ToString(db.Tbldaytiming.Where(x => x.IsDeleted == 0).Select(x => x.StartTime).FirstOrDefault());
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                string startDate = data.fromDate;
                //string endDate = data.toDate; // for selecting multiple day
                string endDate = data.fromDate;
                double dayDifference = Convert.ToDateTime(endDate).Subtract(Convert.ToDateTime(startDate)).TotalDays;
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }

                //var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                foreach (var machRow in machineData)
                {
                    int machieID = machRow.MachineId;
                    for (int i = 0; i <= dayDifference; i++)
                    {
                        DateTime temp = Convert.ToDateTime(startDate).AddDays(i);
                        string correctedDate = temp.ToString("yyyy-MM-dd");
                        var unWODet = db.Tblunasignedwo.Where(x => x.Isupdate == 0 && x.SendApprove == 0 && x.Machineid == machieID && x.Correcteddate == correctedDate).ToList();
                        if (unWODet.Count == 0)
                        {
                            string dayStart = temp.ToString("yyyy-MM-dd");
                            //var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == machieID && x.CorrectedDate == dayStart && x.Status == 2 && x.IsWorkInProgress == 1).Select(x => new { x.Date, x.Time, x.CorrectedDate }).ToList();
                            var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == machieID && x.CorrectedDate == dayStart && (x.Status == 2 || x.Status == 1) && (x.IsWorkInProgress == 1 || x.IsWorkInProgress == 0)).Select(x => new { x.Date, x.Time, x.CorrectedDate }).OrderBy(x => x.Date).ToList();
                            if (hmiData.Count > 0)
                            {
                                for (int j = 0; j <= hmiData.Count(); j++)
                                {
                                    SendHMIDetails objSendHMIDetails = new SendHMIDetails();
                                    if (j == 0) // start time row
                                    {
                                        dayStart = dayStart + " " + dayTimingStartTime;
                                        string dayStartTbl = Convert.ToString(hmiData[j].Date);
                                        double getMinutes = Convert.ToDateTime(dayStartTbl).Subtract(Convert.ToDateTime(dayStart)).TotalMinutes;
                                        if (dayStart == dayStartTbl)
                                        {
                                            // nothing to do
                                        }
                                        else if (getMinutes >= 1)
                                        {
                                            objSendHMIDetails.startDateTime = dayStart;
                                            objSendHMIDetails.endDateTime = dayStartTbl;
                                            objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
                                            objSendHMIDetails.machineId = machieID;
                                        }
                                    }
                                    else if (j == hmiData.Count())// ending time row
                                    {
                                        string previousEndTime = Convert.ToString(hmiData[j - 1].Time);
                                        string tempStartTime = Convert.ToDateTime(startDate).AddDays(1).ToString("yyyy-MM-dd") + " " + dayTimingStartTime;
                                        string presentStartTime = Convert.ToDateTime(tempStartTime).AddSeconds(-1).ToString("yyyy-MM-dd HH:mm:ss");
                                        double getMinutes = Convert.ToDateTime(presentStartTime).Subtract(Convert.ToDateTime(previousEndTime)).TotalMinutes;
                                        if (presentStartTime == previousEndTime)
                                        {
                                            // nothing to do
                                        }
                                        else if (getMinutes >= 1)
                                        {
                                            objSendHMIDetails.startDateTime = previousEndTime;
                                            objSendHMIDetails.endDateTime = presentStartTime;
                                            objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
                                            objSendHMIDetails.machineId = machieID;
                                        }
                                    }
                                    else if (j > 0)
                                    {
                                        string previousEndTime = Convert.ToString(hmiData[j - 1].Time);
                                        string presentStartTime = Convert.ToString(hmiData[j].Date);
                                        double getMinutes = Convert.ToDateTime(presentStartTime).Subtract(Convert.ToDateTime(previousEndTime)).TotalMinutes;
                                        if (presentStartTime == previousEndTime)
                                        {
                                            // nothing to do
                                        }
                                        else if (getMinutes >= 1)
                                        {
                                            objSendHMIDetails.startDateTime = previousEndTime;
                                            objSendHMIDetails.endDateTime = presentStartTime;
                                            objSendHMIDetails.correctedDate = hmiData[i].CorrectedDate;
                                            objSendHMIDetails.machineId = machieID;
                                        }
                                    }
                                    if (objSendHMIDetails.machineId != 0)
                                    {
                                        ret = InsertToUnAsignedWO(objSendHMIDetails);
                                    }
                                }
                            }
                            else if (hmiData.Count == 0) // if there is no record in the hmi table
                            {
                                SendHMIDetails objSendHMIDetails = new SendHMIDetails();
                                DateTime startTime = Convert.ToDateTime(dayStart + " " + dayTimingStartTime);
                                string dayEnd = temp.AddDays(1).ToString("yyyy-MM-dd");
                                DateTime endTime = Convert.ToDateTime(dayEnd + " " + dayTimingStartTime).AddSeconds(-1);
                                objSendHMIDetails.startDateTime = startTime.ToString("yyyy-MM-dd HH:mm:ss");
                                objSendHMIDetails.endDateTime = endTime.ToString("yyyy-MM-dd HH:mm:ss");
                                objSendHMIDetails.correctedDate = temp.ToString("yyyy-MM-dd");
                                objSendHMIDetails.machineId = machieID;
                                ret = InsertToUnAsignedWO(objSendHMIDetails);
                            }
                        }
                        else if (unWODet.Count > 0)
                        {
                            //foreach (var deleteRow in unWODet)// delete previous data and insert new
                            //{
                            //SendHMIDetails objSendHMIDetails = new SendHMIDetails();
                            //objSendHMIDetails.startDateTime = deleteRow.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSendHMIDetails.endDateTime = deleteRow.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSendHMIDetails.correctedDate = deleteRow.Correcteddate;
                            //objSendHMIDetails.machineId = Convert.ToInt32(deleteRow.Machineid);
                            //db.Tblunasignedwo.Remove(deleteRow);
                            //db.SaveChanges();
                            //InsertToUnAsignedWO(objSendHMIDetails);
                            //}
                            //ret = false;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        public bool InsertToUnAsignedWO(SendHMIDetails data)
        {
            bool ret = false;
            int updateLevel = 1;
            try
            {
                var getMailIdsLevel = new List<TblTcfApprovedMaster>();

                var machData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId).Select(x => new { x.CellId, x.ShopId }).FirstOrDefault();
                int cellId = Convert.ToInt32(machData.CellId);
                int shopId = Convert.ToInt32(machData.ShopId);

                getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 2 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
                if (getMailIdsLevel.Count == 0)
                {
                    getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 2 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                }
                foreach (var rowMail in getMailIdsLevel)
                {
                    if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                    {
                        updateLevel = 2;
                    }
                }


                DateTime storeDate = DateTime.Now;
                Tblunasignedwo addRow = new Tblunasignedwo();
                addRow.Starttime = Convert.ToDateTime(data.startDateTime);
                addRow.Endtime = Convert.ToDateTime(data.endDateTime);
                addRow.Pestarttime = Convert.ToDateTime(data.startDateTime);
                addRow.Isdeleted = 0;
                addRow.Createdby = 1;// change as per user login
                addRow.Createdon = Convert.ToString(storeDate);
                addRow.Correcteddate = data.correctedDate;
                addRow.Machineid = data.machineId;
                addRow.UpdateLevel = updateLevel;
                db.Tblunasignedwo.Add(addRow);
                db.SaveChanges();
                ret = true;
            }
            catch (Exception ex)
            {
                ret = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        //Set operatorid
        public CommonResponse SetOperatorId(OperatorDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var unAsignedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == data.uawoid).FirstOrDefault();
                if (unAsignedData != null)
                {
                    unAsignedData.Operatorid = Convert.ToString(data.operatorId);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //set shift
        public CommonResponse SetShift(SetShift data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var unAsignedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == data.uawoid).FirstOrDefault();
                if (unAsignedData != null)
                {
                    unAsignedData.Shiftid = data.shift;
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Split duration

        #region Old Split Dutaion       
        //public CommonResponse SplitDuration(int uawoid)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        var unAssignedWodet = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uawoid).FirstOrDefault();
        //        if (unAssignedWodet != null)
        //        {
        //            string endTime = unAssignedWodet.Endtime.ToString("HH:mm:ss");
        //            string[] splitEndTime = endTime.Split(':');
        //            SplitDuration objSplitDuration = new SplitDuration();
        //            objSplitDuration.startTime = unAssignedWodet.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
        //            objSplitDuration.endDate = unAssignedWodet.Endtime.ToString("yyyy-MM-dd");
        //            objSplitDuration.endHour = splitEndTime[0];
        //            objSplitDuration.endMinute = splitEndTime[1];
        //            objSplitDuration.endSecond = splitEndTime[2];
        //            objSplitDuration.uaWOId = unAssignedWodet.Uawoid;
        //            obj.isTure = true;
        //            obj.response = objSplitDuration;
        //        }
        //        else
        //        {
        //            obj.isTure = false;
        //            obj.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}


        //Compare The Duration With End time
        //public CommonResponse ComapreEndDuration(EntityModels.CompareDuration data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        DateTime endDateTime = Convert.ToDateTime(data.endDate + " " + data.endHour + ":" + data.endMinute + ":" + data.endSecond);
        //        int uaWoId = data.uaWOId;
        //        var unAsiggnedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoId).FirstOrDefault();
        //        if (unAsiggnedData != null)
        //        {
        //            DateTime endTime = unAsiggnedData.Endtime;
        //            DateTime startTime = unAsiggnedData.Starttime;
        //            if (endDateTime <= endTime && endDateTime >= startTime)
        //            {
        //                DateTime nowDateTime = DateTime.Now;
        //                unAsiggnedData.Endtime = endDateTime;
        //                db.SaveChanges();
        //                Tblunasignedwo addRow = new Tblunasignedwo();
        //                addRow.Starttime = endDateTime;
        //                addRow.Endtime = endTime;
        //                addRow.Isdeleted = 0;
        //                addRow.Createdby = 1;// change as per user login
        //                addRow.Createdon = nowDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                addRow.Correcteddate = nowDateTime.ToString("yyyy-MM-dd");
        //                addRow.Machineid = unAsiggnedData.Machineid;
        //                db.Tblunasignedwo.Add(addRow);
        //                db.SaveChanges();
        //            }
        //            else
        //            {
        //                obj.isTure = false;
        //                obj.response = "The End Time Must Be With In " + startTime + "-" + endTime;
        //            }
        //        }
        //        else
        //        {
        //            obj.isTure = false;
        //            obj.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        #endregion

        //Split duration
        public CommonResponse SplitDuration(int uawoid)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var unAssignedWodet = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uawoid).FirstOrDefault();
                if (unAssignedWodet != null)
                {
                    string endTime = unAssignedWodet.Endtime.ToString("HH:mm:ss");
                    string[] splitEndTime = endTime.Split(':');
                    SplitDurationList splitDurationList = new SplitDurationList();
                    List<SplitDuration> splitDurations = new List<SplitDuration>();
                    SplitDuration objSplitDuration = new SplitDuration();
                    objSplitDuration.startTime = unAssignedWodet.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                    objSplitDuration.endTime = unAssignedWodet.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                    //objSplitDuration.endHour = splitEndTime[0];
                    //objSplitDuration.endMinute = splitEndTime[1];
                    //objSplitDuration.endSecond = splitEndTime[2];
                    objSplitDuration.uaWOId = unAssignedWodet.Uawoid;
                    splitDurations.Add(objSplitDuration);
                    splitDurationList.uaWOIds = unAssignedWodet.Uawoid.ToString();
                    splitDurationList.listSplitDuration = splitDurations;
                    obj.isTure = true;
                    obj.response = splitDurationList;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Compare The Duration With End time      
        public CommonResponse ComapreEndDuration(CompareDuration data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int uaWoId = data.uaWOId;
                string allIds = data.uaWOIdS;
                //string[] idList = allIds.Split(',');
                var uaWoData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoId).FirstOrDefault();
                if (uaWoData != null)
                {
                    DateTime endDateTime = Convert.ToDateTime(data.endTime);
                    DateTime endDateTimePrv = Convert.ToDateTime(endDateTime.AddDays(-1));
                    DateTime endTime = uaWoData.Endtime;
                    DateTime startTime = uaWoData.Starttime;
                    if (endDateTime <= endTime && endDateTime >= startTime)
                    {
                        bool check = ValidatePrvEndTime(allIds, endDateTime);
                        if (check)
                        {
                            obj.isTure = false;
                            obj.response = endTime + " This Time Already Exist";
                        }
                        else
                        {
                            string nowDate = uaWoData.Correcteddate;
                            uaWoData.Endtime = endDateTime;
                            db.SaveChanges();
                            Tblunasignedwo addRow = new Tblunasignedwo();
                            addRow.Starttime = endDateTime;
                            addRow.Endtime = endTime;
                            addRow.Isdeleted = 0;
                            addRow.Createdby = 1;// change as per user login
                            addRow.Createdon = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            addRow.Correcteddate = nowDate;
                            addRow.Machineid = uaWoData.Machineid;
                            addRow.IsSplitDuration = uaWoId;
                            addRow.UpdateLevel = uaWoData.UpdateLevel;
                            db.Tblunasignedwo.Add(addRow);
                            db.SaveChanges();
                            allIds += "," + addRow.Uawoid;
                            SplitDurationList objSplitDurationList = new SplitDurationList();
                            objSplitDurationList = GetTheSpliDurationList(allIds);
                            //SplitDuration objSplitDuration = new SplitDuration();
                            //objSplitDuration.startTime = addRow.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSplitDuration.endTime = addRow.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSplitDuration.uaWOId = addRow.Uawoid;
                            //objSplitDuration.uaWOIds=data.uaWOIds+

                            obj.isTure = true;
                            obj.response = objSplitDurationList;
                        }
                    }
                    else if (endDateTimePrv <= endTime && endDateTimePrv >= startTime)
                    {
                        bool check = ValidatePrvEndTime(allIds, endDateTimePrv);
                        if (check)
                        {
                            obj.isTure = false;
                            obj.response = endTime + " This Time Already Exist";
                        }
                        else
                        {
                            string nowDate = uaWoData.Correcteddate;
                            uaWoData.Endtime = endDateTimePrv;
                            db.SaveChanges();
                            Tblunasignedwo addRow = new Tblunasignedwo();
                            addRow.Starttime = endDateTimePrv;
                            addRow.Endtime = endTime;
                            addRow.Isdeleted = 0;
                            addRow.Createdby = 1;// change as per user login
                            addRow.Createdon = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            addRow.Correcteddate = nowDate;
                            addRow.Machineid = uaWoData.Machineid;
                            addRow.IsSplitDuration = uaWoId;
                            addRow.UpdateLevel = uaWoData.UpdateLevel;
                            db.Tblunasignedwo.Add(addRow);
                            db.SaveChanges();
                            allIds += "," + addRow.Uawoid;
                            SplitDurationList objSplitDurationList = new SplitDurationList();
                            objSplitDurationList = GetTheSpliDurationList(allIds);
                            //SplitDuration objSplitDuration = new SplitDuration();
                            //objSplitDuration.startTime = addRow.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSplitDuration.endTime = addRow.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                            //objSplitDuration.uaWOId = addRow.Uawoid;
                            //objSplitDuration.uaWOIds=data.uaWOIds+

                            obj.isTure = true;
                            obj.response = objSplitDurationList;
                        }
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = "The Time Must Be WithIn " + startTime + "-" + endTime;
                    }
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }

            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public bool ValidatePrvEndTime(string allIds, DateTime endDateTime)
        {
            bool result = false;
            try
            {
                string[] idList = allIds.Split(',');
                for (int i = 0; i < idList.Count(); i++)
                {
                    int id = Convert.ToInt32(idList[i]);
                    var check = db.Tblunasignedwo.Where(x => x.Uawoid == id && x.Endtime == endDateTime).FirstOrDefault();
                    if (check != null)
                    {
                        result = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }


        //getting all the list
        public SplitDurationList GetTheSpliDurationList(string uaWoIds)
        {
            SplitDurationList objSplitDurationList = new SplitDurationList();
            List<SplitDuration> listSplitDuration = new List<SplitDuration>();
            string[] splitId = uaWoIds.Split(',');
            for (int i = 0; i < splitId.Count(); i++)
            {
                int uaWoId = Convert.ToInt32(splitId[i]);
                SplitDuration objSplitDuration = new SplitDuration();
                var uaWoData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoId).FirstOrDefault();
                objSplitDuration.startTime = uaWoData.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                objSplitDuration.endTime = uaWoData.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                objSplitDuration.uaWOId = uaWoData.Uawoid;
                listSplitDuration.Add(objSplitDuration);
            }
            objSplitDurationList.uaWOIds = uaWoIds;
            objSplitDurationList.listSplitDuration = listSplitDuration;
            return objSplitDurationList;
        }

        //Deleteing the duration splited
        public CommonResponse DeleteSplitDuration(int uaWOId, string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string storeIDS = "";
                string[] uaWOIDSList = uaWOIds.Split(',');
                for (int i = 0; i < uaWOIDSList.Count(); i++)
                {
                    int spliId = Convert.ToInt32(uaWOIDSList[i]);
                    if (uaWOId == spliId)
                    {
                        if (i == uaWOIDSList.Count() - 1)// last row deleted
                        {
                            var unAssignedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == spliId).FirstOrDefault();
                            int updateId = Convert.ToInt32(uaWOIDSList[i - 1]);
                            var updateRow = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == updateId).FirstOrDefault();
                            updateRow.Endtime = unAssignedData.Endtime;
                            db.SaveChanges();
                            db.Remove(unAssignedData);
                            db.SaveChanges();
                            //obj.isTure = true;
                            //obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                        else
                        {
                            var unAssignedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == spliId).FirstOrDefault();
                            int updateId = Convert.ToInt32(uaWOIDSList[i + 1]);
                            var updateRow = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == updateId).FirstOrDefault();
                            updateRow.Starttime = unAssignedData.Starttime;
                            db.SaveChanges();
                            db.Remove(unAssignedData);
                            db.SaveChanges();
                            //obj.isTure = true;
                            //obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }

                    }
                    else
                    {
                        storeIDS += Convert.ToString(spliId) + ",";
                    }
                }
                storeIDS = storeIDS.TrimEnd(',');
                SplitDurationList objSplitDurationList = new SplitDurationList();
                objSplitDurationList = GetTheSpliDurationList(storeIDS);
                obj.isTure = true;
                obj.response = objSplitDurationList;
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        //Clear all data
        public CommonResponse ClearAllData(string uaWOIds)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] splitIds = uaWOIds.Split(',');
                int count = splitIds.Count();
                int uaWoId = Convert.ToInt32(splitIds[0]);
                int uaWoIdLast = Convert.ToInt32(splitIds[count - 1]);
                DateTime endDateTime = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoIdLast && x.Isdeleted == 0).Select(x => x.Endtime).FirstOrDefault();
                for (int i = 1; i < count; i++)
                {
                    uaWoIdLast = Convert.ToInt32(splitIds[i]);
                    var unawoRemoveRow = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoIdLast && x.Isdeleted == 0).FirstOrDefault();
                    db.Tblunasignedwo.Remove(unawoRemoveRow);
                    db.SaveChanges();
                }
                var unawoUpdateRow = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoId && x.Isdeleted == 0).FirstOrDefault();
                unawoUpdateRow.Endtime = endDateTime;
                db.SaveChanges();
                obj.isTure = true;
                obj.response = "Split Duration Reverted";
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        //Confirm the split duration
        public CommonResponse ConfirmSpliDuration(ConfirmSplitDuration data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] splitIds = data.uaWOIds.Split(',');
                if (data.flageAcceptReject == 1)
                {
                    obj.isTure = true;
                    obj.response = ResourceResponse.SuccessMessage;
                }
                else if (data.flageAcceptReject == 0)
                {
                    int count = splitIds.Count();
                    int uaWoId = Convert.ToInt32(splitIds[0]);
                    int uaWoIdLast = Convert.ToInt32(splitIds[count - 1]);
                    //int uaWoId = Convert.ToInt32(splitIds[count]);
                    DateTime endDateTime = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoIdLast && x.Isdeleted == 0).Select(x => x.Endtime).FirstOrDefault();
                    for (int i = 1; i < count; i++)
                    {
                        uaWoIdLast = Convert.ToInt32(splitIds[i]);
                        var unawoRemoveRow = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoIdLast && x.Isdeleted == 0).FirstOrDefault();
                        db.Tblunasignedwo.Remove(unawoRemoveRow);
                        db.SaveChanges();
                    }
                    var unawoUpdateRow = db.Tblunasignedwo.Where(x => x.Uawoid == uaWoId && x.Isdeleted == 0).FirstOrDefault();
                    unawoUpdateRow.Endtime = endDateTime;
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Split Duration Reverted";
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        //inserting into new table with start time and endtime
        public bool InsertToUnAsignedWOWithGivenEndTime(SendHMIDetails data)
        {
            bool ret = false;
            int updateLevel = 1;
            try
            {
                var getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 6 && x.IsDeleted == 0).ToList();
                foreach (var rowMail in getMailIdsLevel)
                {
                    if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                    {
                        updateLevel = 2;
                    }
                }

                DateTime storeDate = DateTime.Now;
                Tblunasignedwo addRow = new Tblunasignedwo();
                addRow.Starttime = Convert.ToDateTime(data.startDateTime);
                addRow.Endtime = Convert.ToDateTime(data.endDateTime);
                addRow.Isdeleted = 0;
                addRow.Createdby = 1;// change as per user login
                addRow.Createdon = Convert.ToString(storeDate);
                addRow.Correcteddate = data.correctedDate;
                addRow.Machineid = data.machineId;
                addRow.UpdateLevel = updateLevel;
                db.Tblunasignedwo.Add(addRow);
                db.SaveChanges();
                ret = true;
            }
            catch (Exception ex)
            {
                ret = false;
                log.Error(ex);
                if (ex.InnerException != null)
                {
                    log.Error(ex.InnerException.ToString());
                }
            }
            return ret;
        }


        // Getting the all the data based on iscompleted=0 from tblddl
        public DDLCommonResponse GetDDLList(DDLList data)
        {
            DDLCommonResponse obj = new DDLCommonResponse();
            int machineId = data.machineId;
            int takeValue = data.takeValue;
            int skipeValue = data.skipeValue;
            int uaWoId = data.uaWoID;
            try
            {
                //getting the connection string from app string.json
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;

                int count = 0;
                List<EDDLList> listEDDLList = new List<EDDLList>();
                try
                {
                    DataSet ds = new DataSet();
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    SqlCommand SqlCommand = new SqlCommand("GetDDLList", sqlConnection);
                    SqlCommand.CommandType = CommandType.StoredProcedure;
                    SqlCommand.Parameters.AddWithValue("@takeValue", takeValue);
                    SqlCommand.Parameters.AddWithValue("@skipeValue", skipeValue);
                    if (uaWoId != 0)
                    {
                        machineId = Convert.ToInt32(db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoId).Select(x => x.Machineid).FirstOrDefault());
                    }
                    if (machineId == 0)
                    {
                        count = db.Tblddl.Count();
                        SqlCommand.Parameters.AddWithValue("@workCenter", "");
                    }
                    else
                    {

                        string machineInvNo = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => x.MachineInvNo).FirstOrDefault();
                        count = db.Tblddl.Where(m => m.WorkCenter == machineInvNo).Count();
                        SqlCommand.Parameters.AddWithValue("@workCenter", machineInvNo);
                    }
                    SqlDataAdapter da = new SqlDataAdapter(SqlCommand);
                    da.Fill(ds);

                    //var candList = ds.Tables[0].AsEnumerable().ToList();
                    #region
                    if (ds != null && ds.Tables.Count > 0)//check the dataset empty or not
                    {
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            int slno = skipeValue;
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                try
                                {
                                    slno++;
                                    EDDLList objEDDLList = new EDDLList();
                                    objEDDLList.daysAgeing = (ds.Tables[0].Rows[i]["DaysAgeing"]).ToString();
                                    objEDDLList.ddlId = Convert.ToInt32(ds.Tables[0].Rows[i]["DDLID"]);
                                    objEDDLList.flagRush = (ds.Tables[0].Rows[i]["FlagRushInd"]).ToString();
                                    objEDDLList.madDate = (ds.Tables[0].Rows[i]["MADDate"]).ToString();
                                    objEDDLList.madDateInd = (ds.Tables[0].Rows[i]["MADDateInd"]).ToString();
                                    objEDDLList.opNo = (ds.Tables[0].Rows[i]["OperationNo"]).ToString();
                                    objEDDLList.partNo = (ds.Tables[0].Rows[i]["MaterialDesc"]).ToString();
                                    objEDDLList.project = (ds.Tables[0].Rows[i]["Project"]).ToString();
                                    objEDDLList.splitWO = (ds.Tables[0].Rows[i]["SplitWO"]).ToString();
                                    objEDDLList.targetQty = (ds.Tables[0].Rows[i]["TargetQty"]).ToString();
                                    objEDDLList.woNo = (ds.Tables[0].Rows[i]["WorkOrder"]).ToString();
                                    objEDDLList.slno = slno;
                                    listEDDLList.Add(objEDDLList);

                                }
                                catch (Exception ex)
                                {
                                    obj.isTure = false;
                                    obj.response = ResourceResponse.ExceptionMessage;
                                    obj.count = 0;
                                    log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
                                }
                            }
                        }
                    }

                    obj.isTure = true;
                    obj.response = listEDDLList;
                    obj.count = count;

                }
                catch (Exception ex)
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.ExceptionMessage;
                    obj.count = 0;
                    log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
                }

            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }

            return obj;
        }

        //Validate DDL List
        public CommonResponse ValidateDDLIDS(DDLIds data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                //getting the connection string from app string.json
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                string dbName = configuration.GetSection("MySettings").GetSection("Schema").Value;
                SqlConnection sqlConnection = new SqlConnection(connectionString);


                string ddlIds = data.ddlIds;
                string[] ddlCheck = ddlIds.Split(',');
                int count = ddlCheck.Count();
                int ddlId = Convert.ToInt32(ddlCheck[count - 1]);

                SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblddlBackup]", sqlConnection);
                sqlConnection.Open();
                truncateQuery.ExecuteNonQuery();
                sqlConnection.Close();

                for (int i = 0; i < count; i++)
                {
                    log.Error("Insertion Success " + i);
                    SqlCommand insertQuery = new SqlCommand("INSERT INTO " + dbName + ".[tblddlBackup]([ddlIds]) VALUES(" + Convert.ToInt32(ddlCheck[i]) + ")", sqlConnection);
                    sqlConnection.Open();
                    insertQuery.ExecuteNonQuery();
                    sqlConnection.Close();

                    { log.Error("Insertion Success"); }
                    //sqlConnection.Close();
                    int id = Convert.ToInt32(ddlCheck[i]);
                    var ddlRow = db.Tblddl.Where(x => x.Ddlid == id).FirstOrDefault();
                    if (ddlRow != null)
                    {
                        try
                        {
                            DataTable ds = new DataTable();
                            SqlCommand SqlCommand = new SqlCommand("DDLValidate", sqlConnection);
                            SqlCommand.CommandType = CommandType.StoredProcedure;
                            SqlCommand.Parameters.AddWithValue("@operationNo", ddlRow.OperationNo);
                            SqlCommand.Parameters.AddWithValue("@woNo", ddlRow.WorkOrder);
                            SqlCommand.Parameters.AddWithValue("@partNo", ddlRow.MaterialDesc);
                            SqlDataAdapter da = new SqlDataAdapter(SqlCommand);
                            da.Fill(ds);
                            if (ds.Rows.Count > 0)
                            {
                                obj.isTure = false;
                                obj.errorMsg = "Select Operation No:" + ds.Rows[0][3] + " WorkOrder No:" + ds.Rows[0][2] + "Part No:" + ds.Rows[0][5];
                            }
                            else
                            {
                                obj.isTure = true;
                                obj.response = ResourceResponse.SuccessMessage;
                            }
                        }
                        catch (Exception ex)
                        {
                            obj.isTure = false;
                            obj.response = "Stored Procedure Error";
                            log.Error(ex.ToString()); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                        }
                    }


                    else
                    {
                        obj.isTure = false;
                        obj.response = "There is no data";
                        // obj.response = ResourceResponse.FailureMessage;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = "Main Error";
                // obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        //Setting the workOrder for reeworkorder
        public CommonResponse SetReWorkOrder(SetReWork data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int uawoid = data.uaWOId;
                int isChecked = data.isChecked;
                var unAsignedData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uawoid).FirstOrDefault();
                if (unAsignedData != null)
                {
                    if (isChecked == 1)
                    {
                        unAsignedData.Isworkorder = 1;
                    }
                    else
                    {
                        unAsignedData.Isworkorder = 0;
                    }
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        // Getting the shop wise WorkCenter
        public CommonResponse GetShopWiseWorkCenter(int uaWoID)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                //int machineId = data.machineId;
                int machineId = Convert.ToInt32(db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoID).Select(x => x.Machineid).FirstOrDefault());
                List<EShopWiseWC> listEShopWiseWC = new List<EShopWiseWC>();
                int shopId = Convert.ToInt32(db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => x.ShopId).FirstOrDefault());
                if (shopId != 0)
                {
                    var wcDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId && x.IsNormalWc == 0 && x.ManualWcid == null).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    if (wcDetails.Count > 0)
                    {
                        foreach (var row in wcDetails)
                        {
                            EShopWiseWC objEShopWiseWC = new EShopWiseWC();
                            objEShopWiseWC.machineId = row.MachineId;
                            objEShopWiseWC.workCenterName = row.MachineInvNo;
                            listEShopWiseWC.Add(objEShopWiseWC);
                        }
                        obj.isTure = true;
                        obj.response = listEShopWiseWC;
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.NoItemsFound;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //sending the Work Order from DDL to HMI
        public CommonResponse SendWorkOrders(List<SelectWO> ddlIds)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<StoreToUnAsigned> listStoreHMIDetails = new List<StoreToUnAsigned>();
                if (ddlIds.Count > 0)
                {
                    for (int i = 0; i < ddlIds.Count; i++)
                    {
                        StoreToUnAsigned objStoreHMIDetails = new StoreToUnAsigned();
                        int ddlId = Convert.ToInt32(ddlIds[i].ddlId);
                        int unWOId = Convert.ToInt32(ddlIds[i].uaWOId);
                        var ddlList = db.Tblddl.Where(x => x.Ddlid == ddlId)
                            .Select(x => new { x.WorkOrder, x.MaterialDesc, x.OperationNo, x.Project, x.TargetQty, x.WorkCenter, x.Type }).FirstOrDefault();
                        objStoreHMIDetails.uaWOId = unWOId;
                        objStoreHMIDetails.ddlWorkCenter = ddlList.WorkCenter;
                        objStoreHMIDetails.oprationNo = ddlList.OperationNo;
                        objStoreHMIDetails.partNo = ddlList.MaterialDesc;
                        objStoreHMIDetails.project = ddlList.Project;
                        objStoreHMIDetails.workOrderNo = ddlList.WorkOrder;
                        objStoreHMIDetails.workOrderQty = ddlList.TargetQty;
                        objStoreHMIDetails.prodFai = ddlList.Type;
                        listStoreHMIDetails.Add(objStoreHMIDetails);
                    }
                    bool check = InsertToUnAsignedWOData(listStoreHMIDetails);
                    if (check)
                    {
                        //List<SendToHMIScreen> listSendToHMIScreen = new List<SendToHMIScreen>();
                        //var unAsignedWoData = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && x.SendApprove == 0).ToList();
                        //if (unAsignedWoData.Count > 0)
                        //{
                        //    foreach (var row in unAsignedWoData)
                        //    {
                        //        string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.Machineid).Select(x => x.MachineInvNo).FirstOrDefault();
                        //        SendToHMIScreen objSendToHMIScreen = new SendToHMIScreen();
                        //        objSendToHMIScreen.uaWOId = row.Uawoid;
                        //        objSendToHMIScreen.endTime = Convert.ToString(row.Endtime);
                        //        objSendToHMIScreen.machineName = machineName;
                        //        objSendToHMIScreen.oprationNo = row.OprationNo;
                        //        objSendToHMIScreen.partNo = row.PartNo;
                        //        objSendToHMIScreen.startTime = Convert.ToString(row.Starttime);
                        //        objSendToHMIScreen.project = row.Project;
                        //        objSendToHMIScreen.workOrderNo = row.WorkOrderNo;
                        //        objSendToHMIScreen.workOrderQty = row.WorkOrderQty;
                        //        listSendToHMIScreen.Add(objSendToHMIScreen);
                        //    }
                        //    obj.isTure = true;
                        //    obj.response = listSendToHMIScreen;
                        //}
                        //else
                        //{
                        //    obj.isTure = false;
                        //    obj.response = ResourceResponse.FailureMessage;
                        //}
                        obj.isTure = true;
                        obj.response = ResourceResponse.AddedSuccessMessage;

                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Validate the DDL previous selected or not
        public bool CheckPrvDDL(int ddlIds)
        {
            bool ret = false;
            var ddlDetails = db.Tblddl.Where(x => x.IsDeleted == 0 && x.Ddlid == ddlIds).Select(x => new { x.WorkOrder, x.OperationNo, x.MaterialDesc }).FirstOrDefault();
            if (ddlDetails != null)
            {
                var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.WorkOrderNo == ddlDetails.WorkOrder && x.OprationNo == ddlDetails.OperationNo && x.PartNo == ddlDetails.MaterialDesc).FirstOrDefault();
                if (unWO != null)
                {
                    if (unWO.Isworkinprogress == 0)
                    {
                        ret = true;
                    }
                    else
                    {
                        ret = false;
                    }
                }
                else
                {
                    ret = true;
                }
            }

            return ret;
        }

        // insert to unassigned Work order for the new table 
        public bool InsertToUnAsignedWOData(List<StoreToUnAsigned> data)
        {
            bool check = false;
            try
            {
                if (data.Count > 0)
                {
                    int i = 0;
                    string operatorId = "";
                    string shift = "";
                    foreach (var row in data)
                    {
                        string workOrderNo = "", partNo = "", operationNo = "", processedQty = "";
                        workOrderNo = row.workOrderNo;
                        partNo = row.partNo;
                        operationNo = row.oprationNo;
                        processedQty = GetProcessedQty(workOrderNo, partNo, operationNo);
                        if(processedQty =="")
                        {
                            processedQty = null;
                        }
                        Tblunasignedwo updateRow = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == row.uaWOId).FirstOrDefault();
                        if (i == 0)
                        {
                            operatorId = updateRow.Operatorid;
                            shift = updateRow.Shiftid;
                            updateRow.Ddlworkcenter = row.ddlWorkCenter;
                            updateRow.OprationNo = operationNo;
                            updateRow.PartNo = partNo;
                            updateRow.Project = row.project;
                            updateRow.WorkOrderNo = workOrderNo;
                            updateRow.WorkOrderQty = row.workOrderQty;
                            updateRow.ProcessedQty = processedQty;
                            updateRow.ProdFai = row.prodFai;
                            db.SaveChanges();
                            i++;
                        }
                        else
                        {
                            DateTime storeDate = DateTime.Now;
                            Tblunasignedwo addRow = new Tblunasignedwo();
                            addRow.Starttime = Convert.ToDateTime(updateRow.Starttime);
                            addRow.Pestarttime = Convert.ToDateTime(updateRow.Starttime);
                            addRow.Endtime = Convert.ToDateTime(updateRow.Endtime);
                            addRow.Isdeleted = 0;
                            addRow.Createdby = 1;// change as per user login
                            addRow.Createdon = Convert.ToString(storeDate);
                            addRow.Correcteddate = updateRow.Correcteddate;
                            addRow.Machineid = updateRow.Machineid;
                            addRow.Ddlworkcenter = row.ddlWorkCenter;
                            addRow.OprationNo = operationNo;
                            addRow.PartNo = partNo;
                            addRow.Project = row.project;
                            addRow.WorkOrderNo = workOrderNo;
                            addRow.WorkOrderQty = row.workOrderQty;
                            addRow.Operatorid = operatorId;
                            addRow.Shiftid = shift;
                            addRow.ProcessedQty = processedQty;
                            addRow.ProdFai = row.prodFai;
                            addRow.UpdateLevel = updateRow.UpdateLevel;
                            db.Tblunasignedwo.Add(addRow);
                            db.SaveChanges();
                        }
                    }
                    check = true;
                }
                else
                {
                    check = false;
                }
            }
            catch (Exception ex)
            {
                check = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return check;
        }

        //Get processed QTY
        public string GetProcessedQty(string workOrderNo, string partNo, string operationNo)
        {
            string processQty = "";
            int pQty = 0;
            try
            {
                var unAssignedData = db.Tblunasignedwo.Where(x => x.OprationNo == operationNo && x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.Status == 1 && x.Isworkinprogress == 0 && x.Isupdate == 0).OrderByDescending(x => x.Uawoid).FirstOrDefault();
                if (unAssignedData != null)
                {
                    //if(pQty>0)
                    //{
                    //    pQty+= Convert.ToInt32(unAssignedData.DeliveryQty) + Convert.ToInt32(unAssignedData.ProcessedQty);
                    //}
                    //else
                    //{
                    pQty = Convert.ToInt32(unAssignedData.DeliveryQty) + Convert.ToInt32(unAssignedData.ProcessedQty);
                    //}
                }
                else
                {
                    //var hmiData = db.Tblhmiscreen.Where(x => x.OperationNo == operationNo && x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.Status == 1 && x.IsWorkInProgress == 0).OrderByDescending(x => x.Hmiid).FirstOrDefault();
                    var hmiData = db.Tblhmiscreen.Where(x => x.OperationNo == operationNo && x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.IsWorkInProgress == 0).OrderByDescending(x => x.Hmiid).FirstOrDefault();
                    if (hmiData != null)
                    {
                        pQty = Convert.ToInt32(hmiData.DeliveredQty + hmiData.ProcessQty);
                    }
                }
                processQty = Convert.ToString(pQty);
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return processQty;
        }

        //Job finsh the selected work Order
        public CommonResponse JobFinishWorkOrderDetails(StoreHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                if (data.uaWOId != 0)
                {
                    int id = data.uaWOId;
                    Tblunasignedwo updateRow = db.Tblunasignedwo.Where(x => x.Uawoid == id).FirstOrDefault();
                    if (updateRow != null)
                    {
                        int targetQty = Convert.ToInt32(updateRow.WorkOrderQty);
                        int deliveryQrt = Convert.ToInt32(data.deliveryQty) + Convert.ToInt32(updateRow.ProcessedQty);
                        if (deliveryQrt == targetQty)
                        {
                            updateRow.DeliveryQty = data.deliveryQty;
                            updateRow.Status = 2;
                            updateRow.Isworkinprogress = 1;
                            //updateRow.SendApprove = 1;
                            updateRow.ModifiedOn = DateTime.Now;
                            updateRow.ModifiedBy = 1;// Store by session
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = ResourceResponse.SuccessMessage;
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = "The Delivery Qty Must be Equal to Traget Qty";
                        }
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        // partial finish the selected work order
        public CommonResponse PartialFinishWorkOrderDetails(StoreHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                if (data != null)
                {
                    int id = data.uaWOId;
                    int processedQty = 0;
                    Tblunasignedwo updateRow = db.Tblunasignedwo.Where(x => x.Uawoid == id).FirstOrDefault();
                    if (updateRow != null)
                    {
                        string woqty = updateRow.WorkOrderQty;
                        int targetQty = Convert.ToInt32(woqty);
                        int deliqty = Convert.ToInt32(data.deliveryQty);
                        int procqty = Convert.ToInt32(updateRow.ProcessedQty);
                        string calqty = Convert.ToString(deliqty + procqty);
                        string delqty = calqty;
                        int deliveryQrt = Convert.ToInt32(delqty);
                        if (deliveryQrt < targetQty)
                        {
                            string prqty = updateRow.ProcessedQty;
                            //if (prqty != "")
                            //{
                                processedQty= Convert.ToInt32(prqty);

                                if (updateRow.IsSplit == 0)
                                {
                                    string proqty = updateRow.ProcessedQty + data.deliveryQty;
                                    processedQty = Convert.ToInt32(proqty);
                                }
                            //}
                            if (updateRow.Holdcodeid == null || updateRow.Holdcodeid != 1)
                            {
                                updateRow.DeliveryQty = data.deliveryQty;
                                //updateRow.SendApprove = 1;
                                updateRow.Status = 1;
                                updateRow.Isworkinprogress = 0;
                                updateRow.ModifiedOn = DateTime.Now;
                                updateRow.ModifiedBy = 1;// Store by session
                                updateRow.ProcessedQty = processedQty.ToString();
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = ResourceResponse.SuccessMessage;
                            }
                            else
                            {
                                obj.isTure = false;
                                obj.response = "Release Hold Before Finish";
                            }
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = "The Delivery Qty Must be LessThan Traget Qty";
                        }
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //set the row with blank
        public CommonResponse SetBlank(int unwoid, int value)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                if (unwoid != 0)
                {
                    Tblunasignedwo updateRow = db.Tblunasignedwo.Where(x => x.Uawoid == unwoid).FirstOrDefault();
                    if (updateRow != null)
                    {
                        updateRow.IsBlank = value;
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = ResourceResponse.UpdatedSuccessMessage;
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.NoItemsFound;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Remove WorkOrder
        public CommonResponse RemoveWorkOrder(int uawoid)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                Tblunasignedwo removerow = db.Tblunasignedwo.Find(uawoid);
                Tblunasignedwo addRow = new Tblunasignedwo();
                addRow.Starttime = removerow.Starttime;
                addRow.Endtime = removerow.Endtime;
                addRow.Isdeleted = removerow.Isdeleted;
                addRow.Correcteddate = removerow.Correcteddate;
                addRow.Createdby = removerow.Createdby;
                addRow.Createdon = removerow.Createdon;
                addRow.Machineid = removerow.Machineid;
                db.Tblunasignedwo.Add(addRow);
                db.SaveChanges();
                var removeUAWO = db.Tblunasignedwo.Remove(removerow);
                db.SaveChanges();
                obj.isTure = true;
                obj.response = ResourceResponse.SuccessMessage;
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Getting hold code by Levels
        public CommonResponse GetHoldCodes(HoldCode data)
        {
            CommonResponse obj = new CommonResponse();
            int holdCodeLevel = data.holdCodeLevel;
            int holdCodeId = data.holdCodeId;
            try
            {
                if (holdCodeLevel == 1)
                {
                    List<HoldCode> listHoldCode = new List<HoldCode>();
                    var holdCodeDetails = db.Tblholdcodes.Where(x => x.IsDeleted == 0 && x.HoldCodesLevel == 1).ToList();
                    foreach (var row in holdCodeDetails)
                    {
                        HoldCode objHoldCode = new HoldCode();
                        objHoldCode.holdCodeId = row.HoldCodeId;
                        objHoldCode.holdCodeName = row.HoldCodeDesc;
                        objHoldCode.holdCodeLevel = row.HoldCodesLevel;
                        listHoldCode.Add(objHoldCode);
                    }
                    obj.isTure = true;
                    obj.response = listHoldCode;
                }

                else if (holdCodeLevel == 2)
                {
                    List<HoldCode> listHoldCode = new List<HoldCode>();
                    var holdCodeDetails = db.Tblholdcodes.Where(x => x.IsDeleted == 0 && x.HoldCodesLevel1Id == holdCodeId).ToList();
                    if (holdCodeDetails.Count > 0)
                    {
                        foreach (var row in holdCodeDetails)
                        {
                            HoldCode objHoldCode = new HoldCode();
                            objHoldCode.holdCodeId = row.HoldCodeId;
                            objHoldCode.holdCodeName = row.HoldCodeDesc;
                            objHoldCode.holdCodeLevel = row.HoldCodesLevel;
                            listHoldCode.Add(objHoldCode);
                        }
                        obj.isTure = true;
                        obj.response = listHoldCode;
                    }
                    //else
                    //{
                    //    obj.isTure = true;
                    //    obj.response = "Set Reason";
                    //}                    
                }
                else if (holdCodeLevel == 3)
                {
                    List<HoldCode> listHoldCode = new List<HoldCode>();
                    var holdCodeDetails = db.Tblholdcodes.Where(x => x.IsDeleted == 0 && x.HoldCodesLevel2Id == holdCodeId).ToList();
                    if (holdCodeDetails.Count > 0)
                    {
                        foreach (var row in holdCodeDetails)
                        {
                            HoldCode objHoldCode = new HoldCode();
                            objHoldCode.holdCodeId = row.HoldCodeId;
                            objHoldCode.holdCodeName = row.HoldCodeDesc;
                            objHoldCode.holdCodeLevel = row.HoldCodesLevel;
                            listHoldCode.Add(objHoldCode);
                        }
                        obj.isTure = true;
                        obj.response = listHoldCode;
                    }
                    //else
                    //{
                    //    obj.isTure = true;
                    //    obj.response = "Set Reason";
                    //}
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Set the hold code reason for WO
        public CommonResponse SetHoldCodereason(HoldCodeGet data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int uaWoID = data.uaWOId;
                int holdCodeid = data.holdCodeId;
                string holsCodeReason = db.Tblholdcodes.Where(x => x.IsDeleted == 0 && x.HoldCodeId == holdCodeid).Select(x => x.HoldCodeDesc).FirstOrDefault();
                var unAssignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoID).FirstOrDefault();
                if (unAssignedWO != null)
                {
                    unAssignedWO.Holdcodeid = holdCodeid;
                    unAssignedWO.Holdcodereason = holsCodeReason;
                    unAssignedWO.Status = 2;
                    unAssignedWO.Isworkinprogress = 1;
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Getting Generic code by Levels
        public CommonResponse GetGenericCodes(GenericCode data)
        {
            CommonResponse obj = new CommonResponse();
            int genericCodeLevel = data.genericCodeLevel;
            int genericCodeId = data.genericCodeId;
            try
            {
                if (genericCodeLevel == 1)
                {
                    List<GenericCode> listGenericCode = new List<GenericCode>();
                    var genericCodeDetails = db.Tblgenericworkcodes.Where(x => x.IsDeleted == 0 && x.GwcodesLevel == 1).ToList();
                    foreach (var row in genericCodeDetails)
                    {
                        GenericCode objGenericCode = new GenericCode();
                        objGenericCode.genericCodeId = row.GenericWorkId;
                        objGenericCode.genericCodeName = row.GenericWorkDesc;
                        objGenericCode.genericCodeLevel = row.GwcodesLevel;
                        listGenericCode.Add(objGenericCode);
                    }
                    obj.isTure = true;
                    obj.response = listGenericCode;
                }

                else if (genericCodeLevel == 2)
                {
                    List<GenericCode> listGenericCode = new List<GenericCode>();
                    var genericCodeDetails = db.Tblgenericworkcodes.Where(x => x.IsDeleted == 0 && x.GwcodesLevel1Id == genericCodeId).ToList();
                    if (genericCodeDetails.Count > 0)
                    {
                        foreach (var row in genericCodeDetails)
                        {
                            GenericCode objGenericCode = new GenericCode();
                            objGenericCode.genericCodeId = row.GenericWorkId;
                            objGenericCode.genericCodeName = row.GenericWorkDesc;
                            objGenericCode.genericCodeLevel = row.GwcodesLevel;
                            listGenericCode.Add(objGenericCode);
                        }
                        obj.isTure = true;
                        obj.response = listGenericCode;
                    }
                    //else
                    //{
                    //    obj.isTure = true;
                    //    obj.response = "Set Reason";
                    //}                    
                }
                else if (genericCodeLevel == 3)
                {
                    List<GenericCode> listHoldCode = new List<GenericCode>();
                    var genericCodeDetails = db.Tblgenericworkcodes.Where(x => x.IsDeleted == 0 && x.GwcodesLevel2Id == genericCodeId).ToList();
                    if (genericCodeDetails.Count > 0)
                    {
                        foreach (var row in genericCodeDetails)
                        {
                            GenericCode objGenericCode = new GenericCode();
                            objGenericCode.genericCodeId = row.GenericWorkId;
                            objGenericCode.genericCodeName = row.GenericWorkDesc;
                            objGenericCode.genericCodeLevel = row.GwcodesLevel;
                            listHoldCode.Add(objGenericCode);
                        }
                        obj.isTure = true;
                        obj.response = listHoldCode;
                    }
                    //else
                    //{
                    //    obj.isTure = true;
                    //    obj.response = "Set Reason";
                    //}
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Set the Generic code reason for WO
        public CommonResponse SetGenericCodereason(GenericCodeGet data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int uaWoID = data.uaWOId;
                int genericCodeid = data.genericCodeId;
                string genericCodeReason = db.Tblgenericworkcodes.Where(x => x.IsDeleted == 0 && x.GenericWorkId == genericCodeid).Select(x => x.GenericWorkDesc).FirstOrDefault();
                var unAssignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWoID).FirstOrDefault();
                if (unAssignedWO != null)
                {
                    unAssignedWO.GenericCodeid = genericCodeid;
                    unAssignedWO.GenericCodereason = genericCodeReason;
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Split WO
        public CommonResponse SplitWorkOrder(SetSplitWork data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int uaWOId = data.uaWOId;
                var unAssignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Uawoid == uaWOId).FirstOrDefault();
                if (unAssignedWO != null)
                {
                    if (data.isChecked == 1)
                    {
                        unAssignedWO.IsSplit = 1;
                    }
                    else if (data.isChecked == 0)
                    {
                        unAssignedWO.IsSplit = 0;
                    }
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //sending all the workorder for approval
        //public CommonResponse SendToApproveAllWODetails(EntityHMIDetails data)
        //{
        //    CommonResponse obj = new CommonResponse();   // get all the details then send by url or by html
        //    try
        //    {
        //        string correctedDate = data.fromDate;
        //        bool result = StoreIntoUnsignedWO(data);
        //        int cellId = data.cellId;
        //        int machId = data.machineId;
        //        int shopId = data.shopiId;
        //        int plantId = data.plantId;
        //        var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        if (data.machineId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.shopiId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        foreach (var machineRow in machineData)
        //        {
        //            var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 0 && x.Acceptreject == 0 && x.Isupdate == 0 && x.WorkOrderNo != null).ToList();
        //            if (unSignedWO.Count > 0)
        //            {
        //                int sl = 1;
        //                // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");
        //                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate.html");
        //                string htmlStr = File.ReadAllText(reader);

        //                string corredtedDate = "";
        //                //string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";

        //                String[] seperator = { "{{WOStart}}" };
        //                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

        //                var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
        //                htmlStr = htmlStr.Replace("{{WOStart}}", "");
        //                htmlStr = htmlStr.Replace("{{WOEnd}}", "");
        //                foreach (var row in unSignedWO)
        //                {
        //                    //string status = "";
        //                    corredtedDate = row.Correcteddate;
        //                    row.SendApprove = 1;
        //                    db.SaveChanges();
        //                    //if (row.ReWork != null)
        //                    //{
        //                    //    if (row.Status == 2 && row.Isworkinprogress == 1)
        //                    //    {
        //                    //        status = "ReWorkOrder With Job Finish";
        //                    //    }
        //                    //    else if (row.Status == 1 && row.Isworkinprogress == 0)
        //                    //    {
        //                    //        status = "ReWorkOrder With Partial Finish";
        //                    //    }
        //                    //    else
        //                    //    {
        //                    //        status = "ReWorkOrder";
        //                    //    }
        //                    //}
        //                    //else if (row.Status == 1 && row.Isworkinprogress == 0)
        //                    //{
        //                    //    status = "Partial Finish";
        //                    //}
        //                    //else if (row.Status == 2 && row.Isworkinprogress == 1)
        //                    //{
        //                    //    status = "Job Finish";
        //                    //}


        //                    String slno = Convert.ToString(sl);
        //                    int mchId = Convert.ToInt32(row.Machineid);
        //                    String operatorId = Convert.ToString(row.Operatorid);
        //                    String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
        //                    htmlStr = htmlStr.Replace("{{slno}}", slno);
        //                    htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
        //                    htmlStr = htmlStr.Replace("{{Shift}}", row.Shiftid);
        //                    htmlStr = htmlStr.Replace("{{StartTime}}", row.Starttime.ToString("yyyy-MM-dd HH:mm:ss"));
        //                    htmlStr = htmlStr.Replace("{{EndTime}}", row.Endtime.ToString("yyyy-MM-dd HH:mm:ss"));
        //                    htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
        //                    htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
        //                    htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
        //                    htmlStr = htmlStr.Replace("{{OprationNo}}", row.OprationNo);
        //                    htmlStr = htmlStr.Replace("{{Project}}", row.Project);
        //                    htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
        //                    htmlStr = htmlStr.Replace("{{WorkOrderQty}}", row.WorkOrderQty);
        //                    htmlStr = htmlStr.Replace("{{ProcessedQty}}", row.ProcessedQty);
        //                    htmlStr = htmlStr.Replace("{{DeliveryQty}}", row.DeliveryQty);
        //                    //htmlStr = htmlStr.Replace("{{Status}}", status);

        //                    if (unSignedWO.Count == 1)
        //                    {
        //                        htmlStr = htmlStr.Replace("{{WO}}", "");
        //                    }
        //                    else if (sl < unSignedWO.Count)
        //                    {

        //                        htmlStr = htmlStr.Replace("{{WO}}", woHtml);
        //                    }
        //                    sl++;
        //                }

        //                string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURL").Value;
        //                string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURL").Value;


        //                //string rejectSrc = @"C:/TataReport/TCFTemplate/Reject.html?Correcteddate="+ corredtedDate;   
        //                string rejectSrc = rejectUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";
        //                string acceptSrc = acceptUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";

        //                htmlStr = htmlStr.Replace("{{WO}}", "");
        //                htmlStr = htmlStr.Replace("{{userName}}", "Ritesh");
        //                htmlStr = htmlStr.Replace("{{Sname}}", "Saurabh");
        //                //htmlStr = htmlStr.Replace("{{Lurl}}", logo);
        //                htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
        //                htmlStr = htmlStr.Replace("{{urlR}}", rejectSrc);



        //                string toMailID = "monika.ms@srkssolutions.com";
        //                //string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
        //                string ccMailID = "vignesh.pai@srkssolutions.com,aswini.gp@srkssolutions.com";

        //                bool ret = SendMail(htmlStr, toMailID, ccMailID, 1);

        //                //MailMessage mail = new MailMessage();
        //                //mail.To.Add(toMailID);
        //                //mail.CC.Add(ccMailID);
        //                //mail.From = new MailAddress("vignesh.pai@srkssolutions.com");
        //                //mail.Subject = "test mail";
        //                //mail.Body = "" + htmlStr;
        //                //mail.IsBodyHtml = true;


        //                //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlStr, Encoding.UTF8, MediaTypeNames.Text.Html);
        //                //// Create a plain text message for client that don't support HTML
        //                //AlternateView plainView = AlternateView.CreateAlternateViewFromString(Regex.Replace(htmlStr, "<[^>]+?>", string.Empty), Encoding.UTF8, MediaTypeNames.Text.Plain);
        //                //string mediaType = MediaTypeNames.Image.Jpeg;
        //                //LinkedResource img = new LinkedResource(@"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg", mediaType);
        //                //// Make sure you set all these values!!!
        //                //img.ContentId = "EmbeddedContent_1";
        //                //img.ContentType.MediaType = mediaType;
        //                //img.TransferEncoding = TransferEncoding.Base64;
        //                //img.ContentType.Name = img.ContentId;
        //                //img.ContentLink = new Uri("cid:" + img.ContentId);
        //                //LinkedResource img1 = new LinkedResource(@"C:\TataReport\TCFTemplate\approve.Jpeg", mediaType);
        //                //// Make sure you set all these values!!!
        //                //img1.ContentId = "EmbeddedContent_2";
        //                //img1.ContentType.MediaType = mediaType;
        //                //img1.TransferEncoding = TransferEncoding.Base64;
        //                //img1.ContentType.Name = img.ContentId;
        //                //img1.ContentLink = new Uri("cid:" + img1.ContentId);
        //                //LinkedResource img2 = new LinkedResource(@"C:\TataReport\TCFTemplate\reject.Jpeg", mediaType);
        //                //// Make sure you set all these values!!!
        //                //img2.ContentId = "EmbeddedContent_3";
        //                //img2.ContentType.MediaType = mediaType;
        //                //img2.TransferEncoding = TransferEncoding.Base64;
        //                //img2.ContentType.Name = img.ContentId;
        //                //img2.ContentLink = new Uri("cid:" + img2.ContentId);
        //                //htmlView.LinkedResources.Add(img);
        //                //htmlView.LinkedResources.Add(img1);
        //                //htmlView.LinkedResources.Add(img2);
        //                //mail.AlternateViews.Add(plainView);
        //                //mail.AlternateViews.Add(htmlView);


        //                //SmtpClient smtp = new SmtpClient();
        //                //smtp.Host = "smtp.gmail.com";
        //                //smtp.Port = 587;
        //                //smtp.EnableSsl = true;
        //                //smtp.UseDefaultCredentials = false;
        //                //smtp.Credentials = new System.Net.NetworkCredential("vignesh.pai@srkssolutions.com", "vignesh.pai10$");
        //                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
        //                //smtp.Send(mail);
        //                if (ret)
        //                {
        //                    obj.isTure = true;
        //                    obj.response = ResourceResponse.SuccessMessage;
        //                }
        //                else
        //                {
        //                    obj.isTure = false;
        //                    obj.response = ResourceResponse.FailureMessage;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        //Accepting all the work order

        //sending all the workorder for approval

        public CommonResponse SendToApproveAllWODetails(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();   // get all the details then send by url or by html
            try
            {
                string correctedDate = data.fromDate;
                //bool result = StoreIntoUnsignedWO(data);
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                string machName = "";
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
                }

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                string htmlStr = File.ReadAllText(reader);
                String[] seperator = { "{{WOStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{WOStart}}", "");
                htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                string corredtedDate = "";
                int sl = 1;

                foreach (var machineRow in machineData)
                {
                    int i = 0;
                    int machineId = machineRow.MachineId;
                    var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 0 && x.Acceptreject == 0 && x.Isupdate == 0 && x.WorkOrderNo != null && x.Correcteddate == correctedDate && x.Machineid == machineId).OrderBy(m => m.Uawoid).ToList();
                    if (unSignedWO.Count > 0)
                    {

                        // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");                       

                        //string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";

                        foreach (var row in unSignedWO)
                        {
                            //string status = "";
                            corredtedDate = row.Correcteddate;
                            row.SendApprove = 1;
                            row.IsPending = 1;
                            db.SaveChanges();
                            //if (row.ReWork != null)
                            //{
                            //    if (row.Status == 2 && row.Isworkinprogress == 1)
                            //    {
                            //        status = "ReWorkOrder With Job Finish";
                            //    }
                            //    else if (row.Status == 1 && row.Isworkinprogress == 0)
                            //    {
                            //        status = "ReWorkOrder With Partial Finish";
                            //    }
                            //    else
                            //    {
                            //        status = "ReWorkOrder";
                            //    }
                            //}
                            //else if (row.Status == 1 && row.Isworkinprogress == 0)
                            //{
                            //    status = "Partial Finish";
                            //}
                            //else if (row.Status == 2 && row.Isworkinprogress == 1)
                            //{
                            //    status = "Job Finish";
                            //}


                            String slno = Convert.ToString(sl);
                            int mchId = Convert.ToInt32(row.Machineid);
                            String operatorId = Convert.ToString(row.Operatorid);
                            String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{Shift}}", row.Shiftid);
                            htmlStr = htmlStr.Replace("{{StartTime}}", row.Starttime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{EndTime}}", row.Endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.OprationNo);
                            htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                            htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                            htmlStr = htmlStr.Replace("{{WorkOrderQty}}", row.WorkOrderQty);
                            htmlStr = htmlStr.Replace("{{ProcessedQty}}", row.ProcessedQty);
                            htmlStr = htmlStr.Replace("{{DeliveryQty}}", row.DeliveryQty);
                            //htmlStr = htmlStr.Replace("{{Status}}", status);

                            if (unSignedWO.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < unSignedWO.Count)
                            {

                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            sl++;
                        }



                        //MailMessage mail = new MailMessage();
                        //mail.To.Add(toMailID);
                        //mail.CC.Add(ccMailID);
                        //mail.From = new MailAddress("vignesh.pai@srkssolutions.com");
                        //mail.Subject = "test mail";
                        //mail.Body = "" + htmlStr;
                        //mail.IsBodyHtml = true;


                        //AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlStr, Encoding.UTF8, MediaTypeNames.Text.Html);
                        //// Create a plain text message for client that don't support HTML
                        //AlternateView plainView = AlternateView.CreateAlternateViewFromString(Regex.Replace(htmlStr, "<[^>]+?>", string.Empty), Encoding.UTF8, MediaTypeNames.Text.Plain);
                        //string mediaType = MediaTypeNames.Image.Jpeg;
                        //LinkedResource img = new LinkedResource(@"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg", mediaType);
                        //// Make sure you set all these values!!!
                        //img.ContentId = "EmbeddedContent_1";
                        //img.ContentType.MediaType = mediaType;
                        //img.TransferEncoding = TransferEncoding.Base64;
                        //img.ContentType.Name = img.ContentId;
                        //img.ContentLink = new Uri("cid:" + img.ContentId);
                        //LinkedResource img1 = new LinkedResource(@"C:\TataReport\TCFTemplate\approve.Jpeg", mediaType);
                        //// Make sure you set all these values!!!
                        //img1.ContentId = "EmbeddedContent_2";
                        //img1.ContentType.MediaType = mediaType;
                        //img1.TransferEncoding = TransferEncoding.Base64;
                        //img1.ContentType.Name = img.ContentId;
                        //img1.ContentLink = new Uri("cid:" + img1.ContentId);
                        //LinkedResource img2 = new LinkedResource(@"C:\TataReport\TCFTemplate\reject.Jpeg", mediaType);
                        //// Make sure you set all these values!!!
                        //img2.ContentId = "EmbeddedContent_3";
                        //img2.ContentType.MediaType = mediaType;
                        //img2.TransferEncoding = TransferEncoding.Base64;
                        //img2.ContentType.Name = img.ContentId;
                        //img2.ContentLink = new Uri("cid:" + img2.ContentId);
                        //htmlView.LinkedResources.Add(img);
                        //htmlView.LinkedResources.Add(img1);
                        //htmlView.LinkedResources.Add(img2);
                        //mail.AlternateViews.Add(plainView);
                        //mail.AlternateViews.Add(htmlView);


                        //SmtpClient smtp = new SmtpClient();
                        //smtp.Host = "smtp.gmail.com";
                        //smtp.Port = 587;
                        //smtp.EnableSsl = true;
                        //smtp.UseDefaultCredentials = false;
                        //smtp.Credentials = new System.Net.NetworkCredential("vignesh.pai@srkssolutions.com", "vignesh.pai10$");
                        //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                        //smtp.Send(mail);

                    }
                    //sl++;
                }
                htmlStr = htmlStr.Replace(woHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "For 1st Level Approval");

                //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLUnassigned").Value;
                string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLUnassigned").Value;


                //string rejectSrc = @"C:/TataReport/TCFTemplate/Reject.html?Correcteddate="+ corredtedDate;   
                string rejectSrc = rejectUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "&checked=0";
                //string acceptSrc = acceptUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";

                string toName = "";
                //string fromName = "";
                string toMailIds = "";
                string ccMailIds = "";
                //var emailIdCellBase = db.TblEmployee.Where(x => x.Isdeleted == 0 && x.CellId == cellId && x.EmpRole == 9).ToList();
                //foreach (var row in emailIdCellBase)
                //{
                //    if (emailIdCellBase.Count > 1)
                //    {
                //        toName = "All";
                //    }
                //    else
                //    {
                //        toName = row.EmpName;
                //    }
                //    toMailIds += row.EmailId + ",";
                //}


                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.ShopId == shopId).ToList();
                }
                foreach (var row in tcfApproveMail)
                {
                    toMailIds += row.FirstApproverToList + ",";
                    ccMailIds += row.FirstApproverCcList + ",";
                }

                htmlStr = htmlStr.Replace("{{WO}}", "");
                htmlStr = htmlStr.Replace("{{userName}}", toName);
                //htmlStr = htmlStr.Replace("{{Sname}}", "Saurabh");
                //htmlStr = htmlStr.Replace("{{Lurl}}", logo);
                //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);



                //string toMailID = "monika.ms@srkssolutions.com";
                //string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
                //string ccMailID = "vignesh.pai@srkssolutions.com,aswini.gp@srkssolutions.com";

                string toMailID = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                string ccMailID = ccMailIds.Remove(ccMailIds.Length - 1);// removing last comma

                bool ret = SendMail(htmlStr, toMailID, ccMailID, 1, machName);

                if (ret)
                {
                    obj.isTure = true;
                    obj.response = "Sent Mail for Approval";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //public CommonResponse AcceptAllWODetails(EntityHMIDetails data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        int cellId = data.cellId;
        //        int machId = data.machineId;
        //        int shopId = data.shopiId;
        //        int plantId = data.plantId;
        //        string correctedDate = data.fromDate;
        //        var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        if (data.machineId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.shopiId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }

        //        foreach (var machineRow in machineData)
        //        {
        //            int machine = machineRow.MachineId;
        //            var woDetails = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 0 && x.Isupdate == 0 && x.Acceptreject == 0 && x.Correcteddate == correctedDate && x.Machineid == machine).ToList();
        //            if (woDetails.Count > 0)
        //            {
        //                foreach (var row in woDetails)
        //                {
        //                    row.Acceptreject = 1;
        //                    db.SaveChanges();
        //                }

        //                string toMailID = "monika.ms@srkssolutions.com";
        //                //string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
        //                string ccMailID = "vignesh.pai@srkssolutions.com,aswini.gp@srkssolutions.com";

        //                string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
        //                "</head><body><p>Dear Vignesh,</p></br><p><center> The WorkOrder Has Been Accepted</center></p></br><p>Thank you" +
        //                "</p></br><p>Sincerely,</p><p>Dear Vignesh</p></br></body></html>";

        //                bool ret = SendMail(message, toMailID, ccMailID, 0);
        //                if (ret)
        //                {

        //                    obj.isTure = true;
        //                    obj.response = ResourceResponse.SuccessMessage;
        //                    UpdateToReportTables(correctedDate);
        //                }
        //            }
        //            else
        //            {
        //                obj.isTure = false;
        //                obj.response = ResourceResponse.FailureMessage;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        //Getting Reject Reasons


        //Accepting all the work order

        public CommonResponse AcceptAllWODetails(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                string correctedDate = data.fromDate;
                string machName = "";
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
                }
                string toName = "";
                string toMailIds = "";
                string ccMailIds = "";
                bool updateReport = false;
                int appLevel = 0;
                string[] ids = data.id.Split(',');
                foreach (var machineRow in machineData)
                {
                    int machine = machineRow.MachineId;
                    foreach (var idrow in ids)
                    {
                        int uawoid = Convert.ToInt32(idrow);
                        var woDetails = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Isupdate == 0 && (x.Acceptreject == 0 || x.AcceptReject1 == 0) && x.Correcteddate == correctedDate && x.Machineid == machine && x.Uawoid == uawoid).OrderBy(m => m.Uawoid).FirstOrDefault();
                        if (woDetails != null)
                        {
                            //foreach (var row in woDetails)
                            //{
                            if (woDetails.Acceptreject == 0)
                            {
                                woDetails.Acceptreject = 1;
                                woDetails.IsPending = 0;
                                woDetails.ApprovalLevel = 1;
                                db.SaveChanges();
                                if (woDetails.UpdateLevel == 0)
                                {
                                    woDetails.UpdateLevel = 1;
                                    db.SaveChanges();
                                }
                                if (woDetails.UpdateLevel == 1)
                                {
                                    updateReport = true;
                                    appLevel = 1;
                                }
                                else
                                {
                                    appLevel = 2;
                                    log.Error("First Level");
                                }
                            }
                            else if (woDetails.AcceptReject1 == 0)
                            {
                                woDetails.AcceptReject1 = 1;
                                woDetails.IsPending = 0;
                                woDetails.ApprovalLevel = 2;
                                db.SaveChanges();
                                if(woDetails.UpdateLevel == 0)
                                {
                                    woDetails.UpdateLevel = 2;
                                    db.SaveChanges();
                                }
                                if (woDetails.UpdateLevel == 2)
                                {
                                    updateReport = true;
                                }
                                //appLevel = 2;
                                log.Error("Second Level");
                            }
                            //}
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.FailureMessage;
                        }
                    }
                    if (data.unCheckId != "")
                    {
                        string[] unCheckedids = data.unCheckId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedids)
                        {
                            int id = Convert.ToInt32(uncheckedIdRow);
                            var getNoCodeDet = db.Tblunasignedwo.Where(m => m.Uawoid == id).FirstOrDefault();
                            if (getNoCodeDet != null)
                            {
                                getNoCodeDet.IsPending = 1;
                                db.SaveChanges();
                            }
                        }
                    }
                }
                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.ShopId == shopId).ToList();
                }
                foreach (var row in tcfApproveMail)
                {
                    if (appLevel == 1)
                    {
                        toMailIds += row.FirstApproverToList + ",";
                        ccMailIds += row.FirstApproverCcList + ",";
                    }
                    else if (appLevel == 2)
                    {
                        toMailIds += row.SecondApproverToList + ",";
                        ccMailIds += row.SecondApproverCcList + ",";
                    }
                    else if (appLevel == 0)
                    {
                        toMailIds += row.FirstApproverToList + ",";
                        ccMailIds += row.FirstApproverCcList + ",";
                        if (row.SecondApproverToList != "" || row.SecondApproverToList != null)
                        {
                            toMailIds += row.SecondApproverToList + ",";
                            ccMailIds += row.SecondApproverCcList + ",";
                        }
                    }
                }


                string toMailID = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                string ccMailID = ccMailIds.Remove(ccMailIds.Length - 1);

                if (updateReport)
                {
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);
                    String[] seperator = { "{{WOStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{WOStart}}", "");
                    htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                    string corredtedDate = "";
                    int sl = 1;

                    foreach (var machineRow in machineData)
                    {
                        int i = 0;
                        int machineId = machineRow.MachineId;
                        var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Acceptreject == 1 && x.AcceptReject1 == 1 && x.Isupdate == 0 && x.WorkOrderNo != null && x.Correcteddate == correctedDate && x.Machineid == machineId).OrderBy(m => m.Uawoid).ToList();
                        if (unSignedWO.Count > 0)
                        {
                            foreach (var row in unSignedWO)
                            {
                                corredtedDate = row.Correcteddate;
                                //row.SendApprove = 1;
                                //db.SaveChanges();

                                String slno = Convert.ToString(sl);
                                int mchId = Convert.ToInt32(row.Machineid);
                                String operatorId = Convert.ToString(row.Operatorid);
                                String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                                htmlStr = htmlStr.Replace("{{slno}}", slno);
                                htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                                htmlStr = htmlStr.Replace("{{Shift}}", row.Shiftid);
                                htmlStr = htmlStr.Replace("{{StartTime}}", row.Starttime.ToString("yyyy-MM-dd HH:mm:ss"));
                                htmlStr = htmlStr.Replace("{{EndTime}}", row.Endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                                htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                                htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                                htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                                htmlStr = htmlStr.Replace("{{OprationNo}}", row.OprationNo);
                                htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                                htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                                htmlStr = htmlStr.Replace("{{WorkOrderQty}}", row.WorkOrderQty);
                                htmlStr = htmlStr.Replace("{{ProcessedQty}}", row.ProcessedQty);
                                htmlStr = htmlStr.Replace("{{DeliveryQty}}", row.DeliveryQty);

                                if (unSignedWO.Count == 1)
                                {
                                    htmlStr = htmlStr.Replace("{{WO}}", "");
                                }
                                else if (sl < unSignedWO.Count)
                                {

                                    htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                                }
                                else
                                {
                                    htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                                }
                                sl++;
                            }

                        }
                    }
                    htmlStr = htmlStr.Replace(woHtml, "");
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "These Work Orders are Accepted");

                    bool ret = SendMail(htmlStr, toMailID, ccMailID, 2, machName);
                    if (ret)
                    {
                        obj.isTure = true;
                        obj.response = "Sent Mail";
                        UpdateToReportTables(correctedDate,data.machineId);
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }


                    //log.Error("Approved mail");

                    //string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                    //"</head><body><p>Dear " + toName + ",</p></br><p><center> The WorkOrder Has Been Accepted</center></p></br><p>Thank you" +
                    //"</p></body></html>";
                    //bool ret = SendMail(message, toMailID, ccMailID, 0, machName);
                    //if (ret)
                    //{

                    //    obj.isTure = true;
                    //    obj.response = "Data Updated in Reports";
                    //    UpdateToReportTables(correctedDate);
                    //}
                    //else
                    //{
                    //    obj.isTure = false;
                    //    obj.response = ResourceResponse.FailureMessage;
                    //}
                }
                else
                {

                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);
                    String[] seperator = { "{{WOStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{WOStart}}", "");
                    htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                    string corredtedDate = "";
                    int sl = 1;

                    foreach (var machineRow in machineData)
                    {
                        int i = 0;
                        int machineId = machineRow.MachineId;
                        var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Acceptreject == 1 && x.AcceptReject1 == 0 && x.Isupdate == 0 && x.WorkOrderNo != null && x.Correcteddate == correctedDate && x.Machineid == machineId).OrderBy(m => m.Uawoid).ToList();
                        if (unSignedWO.Count > 0)
                        {
                            foreach (var row in unSignedWO)
                            {
                                corredtedDate = row.Correcteddate;
                                //row.SendApprove = 1;
                                //db.SaveChanges();

                                String slno = Convert.ToString(sl);
                                int mchId = Convert.ToInt32(row.Machineid);
                                String operatorId = Convert.ToString(row.Operatorid);
                                String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                                htmlStr = htmlStr.Replace("{{slno}}", slno);
                                htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                                htmlStr = htmlStr.Replace("{{Shift}}", row.Shiftid);
                                htmlStr = htmlStr.Replace("{{StartTime}}", row.Starttime.ToString("yyyy-MM-dd HH:mm:ss"));
                                htmlStr = htmlStr.Replace("{{EndTime}}", row.Endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                                htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                                htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                                htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                                htmlStr = htmlStr.Replace("{{OprationNo}}", row.OprationNo);
                                htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                                htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                                htmlStr = htmlStr.Replace("{{WorkOrderQty}}", row.WorkOrderQty);
                                htmlStr = htmlStr.Replace("{{ProcessedQty}}", row.ProcessedQty);
                                htmlStr = htmlStr.Replace("{{DeliveryQty}}", row.DeliveryQty);

                                if (unSignedWO.Count == 1)
                                {
                                    htmlStr = htmlStr.Replace("{{WO}}", "");
                                }
                                else if (sl < unSignedWO.Count)
                                {

                                    htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                                }
                                else
                                {
                                    htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                                }
                                sl++;
                            }

                        }
                    }
                    htmlStr = htmlStr.Replace(woHtml, "");
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "For 2nd Level Approval");

                    //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLUnassigned").Value;
                    string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLUnassigned").Value;

                    string rejectSrc = rejectUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "&checked=" + data.id + "";
                    //string acceptSrc = acceptUrl + "correctedDate=" + corredtedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";


                    htmlStr = htmlStr.Replace("{{WO}}", "");
                    htmlStr = htmlStr.Replace("{{userName}}", "All");
                    // htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                    htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);


                    bool ret = SendMail(htmlStr, toMailID, ccMailID, 1, machName);
                    if (ret)
                    {
                        log.Error("Seconds Approved mail sent");
                        obj.isTure = true;
                        obj.response = "Sent Mail for Second level Approval";
                        //UpdateToReportTables(correctedDate);
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponse GetRejectReason()
        {
            CommonResponse obj = new CommonResponse();
            List<RejectReasonData> listRejectReasonData = new List<RejectReasonData>();
            var rejectReassondata = db.Tblrejectreason.Where(x => x.IsDeleted == 0 && x.IsTcf == 1).ToList();
            if (rejectReassondata.Count > 0)
            {
                foreach (var row in rejectReassondata)
                {
                    RejectReasonData objRejectReasonData = new RejectReasonData();
                    objRejectReasonData.reasonId = row.Rid;
                    objRejectReasonData.reasonName = row.RejectName;
                    listRejectReasonData.Add(objRejectReasonData);
                }
                obj.isTure = true;
                obj.response = listRejectReasonData;
            }
            return obj;
        }

        ////Rejecting all the work order
        //public CommonResponse RejectAllWODetails(RejectReasonStore data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        string correctedDate = data.correctedDate;
        //        int reasonId = data.reassonId;
        //        int cellId = data.cellId;
        //        int machId = data.machineId;
        //        int shopId = data.shopiId;
        //        int plantId = data.plantId;
        //        var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        if (data.machineId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.shopiId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
        //        }

        //        foreach (var machineRow in machineData)
        //        {
        //            int machine = machineRow.MachineId;
        //            var woDetails = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 0 && x.Isupdate == 0 && x.Acceptreject == 0 && x.Correcteddate == correctedDate && x.Machineid == machine).ToList();
        //            if (woDetails.Count > 0)
        //            {
        //                foreach (var row in woDetails)
        //                {
        //                    row.Acceptreject = 2;
        //                    row.Rejectreasonid = reasonId;
        //                    db.SaveChanges();
        //                }

        //                string toMailID = "monika.ms@srkssolutions.com";
        //                //string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
        //                string ccMailID = "vignesh.pai@srkssolutions.com,aswini.gp@srkssolutions.com";

        //                string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
        //               "</head><body><p>Dear Vignesh,</p></br><p><center> The WorkOrder Has Been Rejected</center></p></br><p>Thank you" +
        //               "</p></br><p>Sincerely,</p><p>Dear Vignesh</p></br></body></html>";

        //                bool ret = SendMail(message, toMailID, ccMailID, 0);

        //                obj.isTure = true;
        //                obj.response = ResourceResponse.SuccessMessage;
        //            }
        //            else
        //            {
        //                obj.isTure = false;
        //                obj.response = ResourceResponse.FailureMessage;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        //Accept Reject Unassigned WO table disaply


        //Rejecting all the work order

        public CommonResponse RejectAllWODetails(RejectReasonStore data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string correctedDate = data.correctedDate;
                int reasonId = data.reassonId;
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                string machName = "";
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "UnAssigned WO " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
                }
                string toName = "";
                string toMailIds = "";
                string ccMailIds = "";
                int checkMail = 0;
                string[] ids = data.id.Split(',');
                foreach (var machineRow in machineData)
                {
                    int machine = machineRow.MachineId;
                    foreach (var idrow in ids)
                    {
                        int uawoid = Convert.ToInt32(idrow);
                        var woDetails = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Isupdate == 0 && (x.Acceptreject == 0 || x.AcceptReject1 == 0) && x.Correcteddate == correctedDate && x.Machineid == machine && x.Uawoid == uawoid).FirstOrDefault();
                        if (woDetails != null)
                        {
                            //foreach (var row in woDetails)
                            //{
                            if (woDetails.Acceptreject == 0)
                            {
                                log.Error("First Level");
                                woDetails.Acceptreject = 2;
                                woDetails.IsPending = 0;
                                woDetails.Rejectreasonid = reasonId;
                                woDetails.ApprovalLevel = 1;
                                db.SaveChanges();
                                checkMail = 1;
                                log.Error("First Level" + checkMail);
                            }
                         
                            else if (woDetails.AcceptReject1 == 0)
                            {
                                log.Error("second Level");
                                woDetails.AcceptReject1 = 2;
                                woDetails.IsPending = 0;
                                woDetails.RejectReason1 = reasonId;
                                woDetails.ApprovalLevel = 2;
                                db.SaveChanges();
                                checkMail = 2;
                                log.Error("Second Level" + checkMail);
                            }

                            // }
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.FailureMessage;
                        }
                    }
                    if (data.unCheckId != "")
                    {
                        string[] unCheckedids = data.unCheckId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedids)
                        {
                            int id = Convert.ToInt32(uncheckedIdRow);
                            var getNoCodeDet = db.Tblunasignedwo.Where(m => m.Uawoid == id).FirstOrDefault();
                            if (getNoCodeDet != null)
                            {
                                getNoCodeDet.IsPending = 1;
                                db.SaveChanges();
                            }
                        }
                    }
                }
                log.Error("CheckMail value");
                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 2 && x.ShopId == shopId).ToList();
                }
                foreach (var row in tcfApproveMail)
                {
                    if (checkMail == 1)
                    {
                        toMailIds += row.FirstApproverToList + ",";
                        ccMailIds += row.FirstApproverCcList + ",";
                    }
                    else if (checkMail == 2)
                    {
                        toMailIds += row.FirstApproverToList + ",";
                        ccMailIds += row.FirstApproverCcList + ",";
                        if (row.SecondApproverToList != "" || row.SecondApproverToList != null)
                        {
                            toMailIds += row.SecondApproverToList + ",";
                            ccMailIds += row.SecondApproverCcList + ",";
                        }
                    }
                }

                log.Error("ToMailid" +toMailIds);
                log.Error("CCMailid" + ccMailIds);
                toMailIds = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);

                string rejectName = db.Tblrejectreason.Where(x => x.IsDeleted == 0 && x.Rid == reasonId).Select(x => x.RejectNameDesc).FirstOrDefault();

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                string htmlStr = File.ReadAllText(reader);
                String[] seperator = { "{{WOStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{WOStart}}", "");
                htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                string corredtedDate = "";
                int sl = 1;

                foreach (var machineRow in machineData)
                {
                    int i = 0;
                    int machineId = machineRow.MachineId;
                    var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.SendApprove == 1 && x.Acceptreject == 2 && x.AcceptReject1 == 2 && x.Isupdate == 0 && x.WorkOrderNo != null && x.Correcteddate == correctedDate && x.Machineid == machineId).OrderBy(m => m.Uawoid).ToList();
                    if (unSignedWO.Count > 0)
                    {
                        foreach (var row in unSignedWO)
                        {
                            corredtedDate = row.Correcteddate;
                            //row.SendApprove = 1;
                            //db.SaveChanges();

                            String slno = Convert.ToString(sl);
                            int mchId = Convert.ToInt32(row.Machineid);
                            String operatorId = Convert.ToString(row.Operatorid);
                            String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{Shift}}", row.Shiftid);
                            htmlStr = htmlStr.Replace("{{StartTime}}", row.Starttime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{EndTime}}", row.Endtime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.OprationNo);
                            htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                            htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                            htmlStr = htmlStr.Replace("{{WorkOrderQty}}", row.WorkOrderQty);
                            htmlStr = htmlStr.Replace("{{ProcessedQty}}", row.ProcessedQty);
                            htmlStr = htmlStr.Replace("{{DeliveryQty}}", row.DeliveryQty);

                            if (unSignedWO.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < unSignedWO.Count)
                            {

                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            sl++;
                        }

                    }
                }
                htmlStr = htmlStr.Replace(woHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "These Work Orders Has Been Rejected For this " + rejectName + ".");



                // string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                //"</head><body><p>Dear All,</p></br><p><center> The WorkOrder Has Been Rejected For this " + rejectName + ".</center></p></br><p>Thank you" +
                //"</p></br></body></html>";

                bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 2, machName);
                if (ret)
                {
                    obj.isTure = true;
                    obj.response = "Sent Reject Mail";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        public CommonResponse AcceptRejectUnWOTable(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] ids = data.id.Split(',');
                List<UnsignedWO> listUnsignedWO = new List<UnsignedWO>();
                string correctedDate = data.fromDate;
                //bool result = StoreIntoUnsignedWO(data);
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                if (data.machineId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                }
                foreach (var machineRow in machineData)
                {
                    int machine = machineRow.MachineId;
                    if (data.id == "")
                    {
                        //var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && x.SendApprove == 1 && x.Correcteddate == correctedDate && x.Machineid == machine && (x.Acceptreject == 0 || x.Acceptreject == 1) && (x.AcceptReject1 == 1 || x.AcceptReject1 == 0)).ToList();
                        var unWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && x.SendApprove == 1 && x.Correcteddate == correctedDate && x.Machineid == machine && x.IsPending == 1).OrderBy(m => m.Uawoid).ToList();
                        if (unWO.Count > 0)
                        {
                            foreach (var row in unWO)
                            {
                                //bool jf = false, pf = false;                            
                                //if (row.Isworkinprogress == 1)
                                //{
                                //    jf = true;
                                //}
                                //if (row.Isworkinprogress == 0)
                                //{
                                //    pf = true;
                                //}                           

                                var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                UnsignedWO objUnsignedWO = new UnsignedWO();
                                objUnsignedWO.uaWOId = row.Uawoid;
                                objUnsignedWO.startDateTime = row.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.endDateTime = row.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                objUnsignedWO.machineName = machineDetais.MachineInvNo;
                                objUnsignedWO.workOrderNo = row.WorkOrderNo;
                                objUnsignedWO.partNo = row.PartNo;
                                objUnsignedWO.oprationNo = row.OprationNo;
                                objUnsignedWO.workOrderQty = row.WorkOrderQty;
                                objUnsignedWO.processedQty = row.ProcessedQty;
                                objUnsignedWO.project = row.Project;
                                objUnsignedWO.prodFai = row.ProdFai;
                                objUnsignedWO.operatorId = row.Operatorid;
                                objUnsignedWO.shift = row.Shiftid;
                                objUnsignedWO.deleveredQty = row.DeliveryQty;
                                objUnsignedWO.deleveredQty = row.DeliveryQty;
                                objUnsignedWO.isWocenter = machineDetais.IsNormalWc;
                                //objUnsignedWO.jf = jf;
                                //objUnsignedWO.pf = pf;
                                listUnsignedWO.Add(objUnsignedWO);
                            }
                            obj.isTure = true;
                            obj.response = listUnsignedWO.OrderBy(x => x.startDateTime);
                        }

                        else if (unWO.Count == 0)
                        {
                            obj.isTure = false;
                            obj.response = "All The Work Order are Accepted";
                        }
                    }
                    else
                    {
                        bool ret = false;
                        foreach (var rowid in ids)
                        {
                            int id = Convert.ToInt32(rowid);
                            var nocdet = db.Tblunasignedwo.Where(x => x.Uawoid == id && x.IsPending == 0).FirstOrDefault();
                            if (nocdet != null)
                            {
                                ret = true;
                            }
                            else
                            {
                                ret = false;
                                break;
                            }
                        }
                        if (ret == true)
                        {
                            foreach (var rowid in ids)
                            {
                                int id = Convert.ToInt32(rowid);
                                var NoCodeDet = db.Tblunasignedwo.Where(x => x.Uawoid == id && x.IsPending == 0 && (x.Acceptreject == 1 || x.AcceptReject1 == 1) && x.ApprovalLevel != 2).FirstOrDefault();
                                if (NoCodeDet != null)
                                {

                                    //foreach (var row in NoCodeDet)
                                    //{
                                    var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == NoCodeDet.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                    UnsignedWO objUnsignedWO = new UnsignedWO();
                                    objUnsignedWO.uaWOId = NoCodeDet.Uawoid;
                                    objUnsignedWO.startDateTime = NoCodeDet.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                    objUnsignedWO.endDateTime = NoCodeDet.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                    objUnsignedWO.machineName = machineDetais.MachineInvNo;
                                    objUnsignedWO.workOrderNo = NoCodeDet.WorkOrderNo;
                                    objUnsignedWO.partNo = NoCodeDet.PartNo;
                                    objUnsignedWO.oprationNo = NoCodeDet.OprationNo;
                                    objUnsignedWO.workOrderQty = NoCodeDet.WorkOrderQty;
                                    objUnsignedWO.processedQty = NoCodeDet.ProcessedQty;
                                    objUnsignedWO.project = NoCodeDet.Project;
                                    objUnsignedWO.prodFai = NoCodeDet.ProdFai;
                                    objUnsignedWO.operatorId = NoCodeDet.Operatorid;
                                    objUnsignedWO.shift = NoCodeDet.Shiftid;
                                    objUnsignedWO.deleveredQty = NoCodeDet.DeliveryQty;
                                    objUnsignedWO.deleveredQty = NoCodeDet.DeliveryQty;
                                    objUnsignedWO.isWocenter = machineDetais.IsNormalWc;
                                    //objUnsignedWO.jf = jf;
                                    //objUnsignedWO.pf = pf;
                                    listUnsignedWO.Add(objUnsignedWO);
                                    //}
                                    obj.isTure = true;
                                    obj.response = listUnsignedWO.OrderBy(x => x.startDateTime);
                                }

                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "All The Work Order are Accepted";
                                }
                            }
                        }
                        else
                        {
                            foreach (var rowid in ids)
                            {
                                int id = Convert.ToInt32(rowid);
                                var NoCodeDet = db.Tblunasignedwo.Where(x => x.Uawoid == id && x.IsPending == 1).FirstOrDefault();
                                if (NoCodeDet != null)
                                {

                                    //foreach (var row in NoCodeDet)
                                    //{
                                    var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == NoCodeDet.Machineid).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                    UnsignedWO objUnsignedWO = new UnsignedWO();
                                    objUnsignedWO.uaWOId = NoCodeDet.Uawoid;
                                    objUnsignedWO.startDateTime = NoCodeDet.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                    objUnsignedWO.endDateTime = NoCodeDet.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                    objUnsignedWO.machineName = machineDetais.MachineInvNo;
                                    objUnsignedWO.workOrderNo = NoCodeDet.WorkOrderNo;
                                    objUnsignedWO.partNo = NoCodeDet.PartNo;
                                    objUnsignedWO.oprationNo = NoCodeDet.OprationNo;
                                    objUnsignedWO.workOrderQty = NoCodeDet.WorkOrderQty;
                                    objUnsignedWO.processedQty = NoCodeDet.ProcessedQty;
                                    objUnsignedWO.project = NoCodeDet.Project;
                                    objUnsignedWO.prodFai = NoCodeDet.ProdFai;
                                    objUnsignedWO.operatorId = NoCodeDet.Operatorid;
                                    objUnsignedWO.shift = NoCodeDet.Shiftid;
                                    objUnsignedWO.deleveredQty = NoCodeDet.DeliveryQty;
                                    objUnsignedWO.deleveredQty = NoCodeDet.DeliveryQty;
                                    objUnsignedWO.isWocenter = machineDetais.IsNormalWc;
                                    //objUnsignedWO.jf = jf;
                                    //objUnsignedWO.pf = pf;
                                    listUnsignedWO.Add(objUnsignedWO);
                                    //}
                                    obj.isTure = true;
                                    obj.response = listUnsignedWO.OrderBy(x => x.startDateTime);
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {

            }
            return obj;
        }

        //getting all the record for updation into table
        public CommonResponse UpdateToReportTables(string correctedDate,int MachineID)
        {
            bool check = false;
            CommonResponse obj = new CommonResponse();
            try
            {
                var unSignedWO = db.Tblunasignedwo.Where(x => x.Isdeleted == 0 && x.Isupdate == 0 && (x.Acceptreject == 1 || x.AcceptReject1 == 1) && x.OprationNo != null && x.WorkOrderNo != null && x.PartNo != null && x.Correcteddate == correctedDate).OrderBy(m => m.Uawoid).ToList();
                foreach (var row in unSignedWO)
                {
                    check = InsertToLiveHMI(row);
                    check = InsertToLoginDetails(row);
                }
                if (check)
                {
                    check = TakeBackupReportData(correctedDate);
                    if (check)
                    {
                        DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
                        DateTime correcteDateTime = Convert.ToDateTime(correctedDate);
                        List<int> machinelist = new List<int>();
                        machinelist.Add(MachineID);
                        Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(correcteDateTime, correcteDateTime,machinelist);  // for WO report updation
                        Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(correcteDateTime, correcteDateTime,machinelist);// for OEE report updation
                        if (reportOEEUpdate.Result == true && reportOEEUpdate.Result == true)
                        {
                            foreach (var row in unSignedWO)
                            {
                                row.Isupdate = 1;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = ResourceResponse.SuccessMessage;
                            }
                        }
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Report Backup Data
        public bool TakeBackupReportData(string correctedDate)
        {
            bool result = false;
            //getting the connection string from app string.json
            string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
            string dbName = configuration.GetSection("MySettings").GetSection("Schema").Value;
            DataTable dt = new DataTable();
            string queryUnAssignedMachine = "select Distinct(Machineid) from " + dbName + ".[tblunasignedwo] where Correcteddate='" + correctedDate + "' and (Acceptreject=1 or Acceptreject1=1) and Isupdate=0";
            SqlDataAdapter sDA = new SqlDataAdapter(queryUnAssignedMachine, connectionString);
            sDA.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DateTime corrDate = Convert.ToDateTime(correctedDate + " " + "00:00:00");
                    int machineId = Convert.ToInt32(dt.Rows[i][0]);
                    var oeeDasjboardVar = db.Tbloeedashboardvariables.Where(x => x.IsDeleted == 0 && x.Wcid == machineId && x.StartDate == corrDate).FirstOrDefault();
                    bool oeeCheck = InsertToOEEDashboardVar(oeeDasjboardVar);
                    log.Error("oeeCheck" + oeeCheck);
                    var woreport = db.Tblworeport.Where(x => x.CorrectedDate == correctedDate && x.MachineId == machineId).ToList();
                    bool woReportCheck = InsertToWoReport(woreport);
                    log.Error("woReportCheck" + woReportCheck);
                    var woLoss = db.Tblhmiscreen.Where(x => x.MachineId == machineId && x.CorrectedDate == correctedDate).Select(x => x.Hmiid).ToList();
                    log.Error("hmiids for Woloss" + woLoss.Count);
                    bool woLossCheck = InsertToWoLoss(woLoss);
                    if (oeeCheck == true)
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
            }
            return result;
        }

        //insert method for oeeDashboardVariable Backup
        public bool InsertToOEEDashboardVar(Tbloeedashboardvariables oeeDasjboardVar)
        {
            bool res = true;
            try
            {
                if (oeeDasjboardVar != null)
                {
                    TbloeedashboardvariablesBackup addRow = new TbloeedashboardvariablesBackup();
                    addRow.Blue = oeeDasjboardVar.Blue;
                    addRow.CellId = oeeDasjboardVar.CellId;
                    addRow.CreatedBy = oeeDasjboardVar.CreatedBy;
                    addRow.CreatedOn = oeeDasjboardVar.CreatedOn;
                    addRow.DownTimeBreakdown = oeeDasjboardVar.DownTimeBreakdown;
                    addRow.EndDate = oeeDasjboardVar.EndDate;
                    addRow.Green = oeeDasjboardVar.Green;
                    addRow.IsDeleted = oeeDasjboardVar.IsDeleted;
                    addRow.Loss1Name = oeeDasjboardVar.Loss1Name;
                    addRow.Loss1Value = oeeDasjboardVar.Loss1Value;
                    addRow.Loss2Name = oeeDasjboardVar.Loss2Name;
                    addRow.Loss2Value = oeeDasjboardVar.Loss2Value;
                    addRow.Loss3Name = oeeDasjboardVar.Loss3Name;
                    addRow.Loss3Value = oeeDasjboardVar.Loss3Value;
                    addRow.Loss4Name = oeeDasjboardVar.Loss4Name;
                    addRow.Loss4Value = oeeDasjboardVar.Loss4Value;
                    addRow.Loss5Name = oeeDasjboardVar.Loss5Name;
                    addRow.Loss5Value = oeeDasjboardVar.Loss5Value;
                    addRow.MinorLosses = oeeDasjboardVar.MinorLosses;
                    //addRow.OeevariablesBackupId = oeeDasjboardVar.OeevariablesId;
                    addRow.PlantId = oeeDasjboardVar.PlantId;
                    addRow.ReWotime = oeeDasjboardVar.ReWotime;
                    addRow.Roalossess = oeeDasjboardVar.Roalossess;
                    addRow.ScrapQtyTime = oeeDasjboardVar.ScrapQtyTime;
                    addRow.SettingTime = oeeDasjboardVar.SettingTime;
                    addRow.ShopId = oeeDasjboardVar.ShopId;
                    addRow.StartDate = oeeDasjboardVar.StartDate;
                    addRow.SummationOfSctvsPp = oeeDasjboardVar.SummationOfSctvsPp;
                    addRow.Wcid = oeeDasjboardVar.Wcid;
                    db.TbloeedashboardvariablesBackup.Add(addRow);
                    db.SaveChanges();

                    db.Tbloeedashboardvariables.Remove(oeeDasjboardVar);
                    db.SaveChanges();
                    res = true;
                }
                else
                {
                    res = true;
                }
            }
            catch (Exception ex)
            {
                res = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return res;
        }

        //insert method for woreport Backup
        public bool InsertToWoReport(List<Tblworeport> woreport)
        {
            bool result = false;
            try
            {
                if (woreport.Count > 0)
                {
                    foreach (var row in woreport)
                    {
                        TblworeportBackup addRow = new TblworeportBackup();
                        addRow.BatchNo = row.BatchNo;
                        addRow.Blue = row.Blue;
                        addRow.Breakdown = row.Breakdown;
                        addRow.CorrectedDate = row.CorrectedDate;
                        addRow.CuttingTime = row.CuttingTime;
                        addRow.DeliveredQty = row.DeliveredQty;
                        addRow.EndTime = row.EndTime;
                        addRow.Hmiid = row.Hmiid;
                        addRow.HoldReason = row.HoldReason;
                        addRow.Idle = row.Idle;
                        addRow.InsertedOn = row.InsertedOn;
                        addRow.IsHold = row.IsHold;
                        addRow.IsMultiWo = row.IsMultiWo;
                        addRow.IsNormalWc = row.IsNormalWc;
                        addRow.IsPf = row.IsPf;
                        addRow.MachineId = row.MachineId;
                        addRow.MinorLoss = row.MinorLoss;
                        addRow.Mrweight = row.Mrweight;
                        addRow.NccuttingTimePerPart = row.NccuttingTimePerPart;
                        addRow.OperatorName = row.OperatorName;
                        addRow.OpNo = row.OpNo;
                        addRow.PartNo = row.PartNo;
                        addRow.Program = row.Program;
                        addRow.RejectedQty = row.RejectedQty;
                        addRow.RejectedReason = row.RejectedReason;
                        addRow.ReWorkTime = row.ReWorkTime;
                        addRow.ScrapQtyTime = row.ScrapQtyTime;
                        addRow.SelfInspection = row.SelfInspection;
                        addRow.SettingTime = row.SettingTime;
                        addRow.Shift = row.Shift;
                        addRow.SplitWo = row.SplitWo;
                        addRow.StartTime = row.StartTime;
                        addRow.SummationOfSctvsPp = row.SummationOfSctvsPp;
                        addRow.TargetQty = row.TargetQty;
                        addRow.TotalNccuttingTime = row.TotalNccuttingTime;
                        addRow.Type = row.Type;
                        addRow.Woefficiency = row.Woefficiency;
                        //addRow.WoreportBackupId = row.WoreportId;
                        addRow.WorkOrderNo = row.WorkOrderNo;
                        db.TblworeportBackup.Add(addRow);
                        db.SaveChanges();

                        db.Tblworeport.Remove(row);
                        db.SaveChanges();

                        result = true;
                    }
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //Insert method for woloss Backup
        public bool InsertToWoLoss(List<int> woLoss)
        {
            bool result = false;
            try
            {
                log.Error("InsertToWoloss try" + woLoss);
                if (woLoss.Count > 0 && woLoss != null)
                {
                    foreach (int hmiid in woLoss)
                    {
                        log.Error("InsertToWoloss" + hmiid);
                        var woLossRow = db.Tblwolossess.Where(x => x.IsDeleted == 0 && x.Hmiid == hmiid).FirstOrDefault();
                        if (woLossRow != null)
                        {
                            TblwolossessBackup addRow = new TblwolossessBackup();
                            addRow.Hmiid = woLossRow.Hmiid;
                            addRow.InsertedOn = woLossRow.InsertedOn;
                            addRow.IsDeleted = woLossRow.IsDeleted;
                            addRow.Level = woLossRow.Level;
                            addRow.LossCodeLevel1Id = woLossRow.LossCodeLevel1Id;
                            addRow.LossCodeLevel1Name = woLossRow.LossCodeLevel1Name;
                            addRow.LossCodeLevel2Id = woLossRow.LossCodeLevel2Id;
                            addRow.LossCodeLevel2Name = woLossRow.LossCodeLevel2Name;
                            addRow.LossDuration = woLossRow.LossDuration;
                            addRow.LossId = woLossRow.LossId;
                            addRow.LossName = woLossRow.LossName;
                            //addRow.WolossesBackupId = woLossRow.WolossesId;
                            db.TblwolossessBackup.Add(addRow);
                            db.SaveChanges();
                            log.Error("InsertToWoloss inserted");
                            db.Tblwolossess.Remove(woLossRow);
                            db.SaveChanges();
                            log.Error("InsertToWoloss Removed");
                            result = true;
                        }
                        else
                        {
                            result = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Error("InsertToWoloss Exception");
                log.Error(ex.ToString());
                if (ex.InnerException.ToString() != null)
                {
                    log.Error(ex.InnerException.ToString());
                }
            }
            return result;
        }

        //Insert to login details table
        public bool InsertToLoginDetails(Tblunasignedwo data)
        {
            bool result = false;
            try
            {
                LoginDetails addRow = new LoginDetails();
                addRow.CorrectedDate = data.Correcteddate;
                addRow.CreatedBy = DateTime.Now;
                addRow.CreatedOn = Convert.ToDateTime(data.Createdon);
                addRow.EndTime = data.Endtime;
                addRow.IsActive = false;
                addRow.IsCompleted = false;
                addRow.IsDeleted = false;
                addRow.MachineId = data.Machineid;
                addRow.StartTime = data.Starttime;
                db.LoginDetails.Add(addRow);
                db.SaveChanges();
                result = true;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //insert method for the report table and related tables
        public bool InsertToLiveHMI(Tblunasignedwo data)
        {
            bool ret = false;
            try
            {
                Tbllivehmiscreen addRow = new Tbllivehmiscreen();
                addRow.CorrectedDate = data.Correcteddate;
                addRow.Date = Convert.ToDateTime(data.Starttime);
                addRow.PestartTime = Convert.ToString(data.Pestarttime);
                addRow.DeliveredQty = Convert.ToInt32(data.DeliveryQty);
                addRow.IsWorkInProgress = 0;
                addRow.MachineId = Convert.ToInt32(data.Machineid);
                addRow.OperationNo = data.OprationNo;
                addRow.OperatorDet = Convert.ToString(data.Operatorid);
                addRow.PartNo = data.PartNo;
                addRow.ProcessQty = Convert.ToInt32(data.ProcessedQty);
                addRow.ProdFai = data.ProdFai;
                addRow.Project = data.Project;
                addRow.Shift = data.Shiftid;
                addRow.Status = data.Status;
                addRow.TargetQty = Convert.ToInt32(data.WorkOrderQty);
                addRow.Time = Convert.ToDateTime(data.Endtime);
                addRow.WorkOrderNo = data.WorkOrderNo;
                addRow.DdlwokrCentre = data.Ddlworkcenter;
                addRow.IsWorkOrder = Convert.ToInt32(data.Isworkorder);
                db.Tbllivehmiscreen.Add(addRow);
                db.SaveChanges();

                int hmiid = addRow.Hmiid;

                var updareLiveRow = db.Tbllivehmiscreen.Where(x => x.Hmiid == hmiid).FirstOrDefault();
                updareLiveRow.Status = data.Status;
                updareLiveRow.IsWorkInProgress = Convert.ToInt32(data.Isworkinprogress);
                updareLiveRow.DdlwokrCentre = data.Ddlworkcenter;
                updareLiveRow.IsWorkOrder = Convert.ToInt32(data.Isworkorder);
                updareLiveRow.SplitWo = Convert.ToString(data.IsSplit);
                updareLiveRow.PestartTime = Convert.ToString(data.Pestarttime);
                db.SaveChanges();

                Tblhmiscreen addRowHmi = new Tblhmiscreen();
                addRowHmi.Hmiid = hmiid;
                addRowHmi.CorrectedDate = data.Correcteddate;
                addRowHmi.Date = Convert.ToDateTime(data.Starttime);
                addRowHmi.DeliveredQty = Convert.ToInt32(data.DeliveryQty);
                addRowHmi.PestartTime = data.Pestarttime;
                addRowHmi.IsWorkInProgress = 0;
                addRowHmi.MachineId = Convert.ToInt32(data.Machineid);
                addRowHmi.OperationNo = data.OprationNo;
                addRowHmi.OperatorDet = Convert.ToString(data.Operatorid);
                addRowHmi.PartNo = data.PartNo;
                addRowHmi.ProcessQty = Convert.ToInt32(data.ProcessedQty);
                addRowHmi.ProdFai = data.ProdFai;
                addRowHmi.Project = data.Project;
                addRowHmi.Shift = data.Shiftid;
                addRowHmi.Status = data.Status;
                addRowHmi.TargetQty = Convert.ToInt32(data.WorkOrderQty);
                addRowHmi.Time = Convert.ToDateTime(data.Endtime);
                addRowHmi.WorkOrderNo = data.WorkOrderNo;
                addRow.DdlwokrCentre = data.Ddlworkcenter;
                db.Tblhmiscreen.Add(addRowHmi);
                db.SaveChanges();

                var updareRow = db.Tblhmiscreen.Where(x => x.Hmiid == hmiid).FirstOrDefault();
                updareRow.Status = data.Status;
                updareRow.IsWorkInProgress = Convert.ToInt32(data.Isworkinprogress);
                updareRow.DdlwokrCentre = data.Ddlworkcenter;
                updareRow.IsWorkOrder = Convert.ToInt32(data.Isworkorder);
                updareRow.PestartTime = data.Pestarttime;
                updareRow.SplitWo = Convert.ToString(data.IsSplit);
                db.SaveChanges();

                if (data.Isworkinprogress == 1 && data.Status == 2)
                {
                    var updateRow = db.Tblddl.Where(x => x.IsDeleted == 0 && x.WorkOrder == data.WorkOrderNo && x.MaterialDesc == data.PartNo && x.OperationNo == data.OprationNo).FirstOrDefault();
                    if (updateRow != null)
                    {
                        updateRow.IsCompleted = 1;
                        db.SaveChanges();
                    }
                }

                ret = true;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        ////method for mailing with image=1 without image=0
        //public bool SendMail(string message, string toList, string ccList, int image)
        //{
        //    bool ret = false;
        //    try
        //    {
        //        if (message != "" && toList != "" && ccList != null)
        //        {

        //            //string toMailID = "monika.ms@srkssolutions.com";
        //            ////string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
        //            //string ccMailID = "vignesh.pai@srkssolutions.com";
        //            string toMailID = toList;
        //            string ccMailID = ccList;
        //            MailMessage mail = new MailMessage();
        //            mail.To.Add(toMailID);
        //            mail.CC.Add(ccMailID);
        //            mail.From = new MailAddress("vignesh.pai@srkssolutions.com");
        //            mail.Subject = "UnAssigned WorkOrder";
        //            mail.Body = "" + message;
        //            mail.IsBodyHtml = true;

        //            if (image == 1)
        //            {
        //                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, MediaTypeNames.Text.Html);
        //                // Create a plain text message for client that don't support HTML
        //                AlternateView plainView = AlternateView.CreateAlternateViewFromString(Regex.Replace(message, "<[^>]+?>", string.Empty), Encoding.UTF8, MediaTypeNames.Text.Plain);
        //                string mediaType = MediaTypeNames.Image.Jpeg;
        //                LinkedResource img = new LinkedResource(@"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg", mediaType);
        //                // Make sure you set all these values!!!
        //                img.ContentId = "EmbeddedContent_1";
        //                img.ContentType.MediaType = mediaType;
        //                img.TransferEncoding = TransferEncoding.Base64;
        //                img.ContentType.Name = img.ContentId;
        //                img.ContentLink = new Uri("cid:" + img.ContentId);
        //                LinkedResource img1 = new LinkedResource(@"C:\TataReport\TCFTemplate\approve.Jpeg", mediaType);
        //                // Make sure you set all these values!!!
        //                img1.ContentId = "EmbeddedContent_2";
        //                img1.ContentType.MediaType = mediaType;
        //                img1.TransferEncoding = TransferEncoding.Base64;
        //                img1.ContentType.Name = img.ContentId;
        //                img1.ContentLink = new Uri("cid:" + img1.ContentId);
        //                LinkedResource img2 = new LinkedResource(@"C:\TataReport\TCFTemplate\reject.Jpeg", mediaType);
        //                // Make sure you set all these values!!!
        //                img2.ContentId = "EmbeddedContent_3";
        //                img2.ContentType.MediaType = mediaType;
        //                img2.TransferEncoding = TransferEncoding.Base64;
        //                img2.ContentType.Name = img.ContentId;
        //                img2.ContentLink = new Uri("cid:" + img2.ContentId);
        //                htmlView.LinkedResources.Add(img);
        //                htmlView.LinkedResources.Add(img1);
        //                htmlView.LinkedResources.Add(img2);
        //                mail.AlternateViews.Add(plainView);
        //                mail.AlternateViews.Add(htmlView);
        //            }


        //            SmtpClient smtp = new SmtpClient();
        //            smtp.Host = "smtp.gmail.com";
        //            smtp.Port = 587;
        //            smtp.EnableSsl = true;
        //            smtp.UseDefaultCredentials = false;
        //            smtp.Credentials = new System.Net.NetworkCredential("vignesh.pai@srkssolutions.com", "vignesh.pai10$");
        //            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
        //            smtp.Send(mail);


        //            ret = true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ret = false;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return ret;
        //}

        //For Generic WorkOrder


        //method for mailing with image=1 without image=0

        public bool SendMail(string message, string toList, string ccList, int image, string subject)
        {
            bool ret = false;
            try
            {
                if (message != "" && toList != "" && ccList != null)
                {

                    //string toMailID = "monika.ms@srkssolutions.com";
                    ////string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
                    //string ccMailID = "vignesh.pai@srkssolutions.com";
                    string toMailID = toList;
                    string ccMailID = ccList;
                    MailMessage mail = new MailMessage();
                    mail.To.Add(toMailID);
                    if (ccMailID != "")
                    {
                        mail.CC.Add(ccMailID);
                    }

                    var smtpConn = db.Smtpdetails.Where(x => x.IsDeleted == true && x.TcfModuleId == 2).FirstOrDefault();
                    string hostName = smtpConn.Host;
                    int port = smtpConn.Port;
                    bool enableSsl = smtpConn.EnableSsl;
                    bool useDefaultCredentials = smtpConn.UseDefaultCredentials;
                    string emailId = smtpConn.EmailId;
                    string password = smtpConn.Password;
                    string fromMail = smtpConn.FromMailId;
                    string connectType = "";
                    if (smtpConn.ConnectType != "" || smtpConn.ConnectType != null)
                    {
                        connectType = smtpConn.ConnectType;//domain
                    }


                    //string fromMail= configuration.GetSection("SMTPConn").GetSection("FromMailID").Value;

                    mail.From = new MailAddress(fromMail);
                    mail.Subject = subject;
                    mail.Body = "" + message;
                    mail.IsBodyHtml = true;

                    if (image == 1)
                    {
                        AlternateView htmlView = AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, MediaTypeNames.Text.Html);
                        // Create a plain text message for client that don't support HTML
                        AlternateView plainView = AlternateView.CreateAlternateViewFromString(Regex.Replace(message, "<[^>]+?>", string.Empty), Encoding.UTF8, MediaTypeNames.Text.Plain);
                        string mediaType = MediaTypeNames.Image.Jpeg;
                        LinkedResource img = new LinkedResource(@"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg", mediaType);
                        // Make sure you set all these values!!!
                        img.ContentId = "EmbeddedContent_1";
                        img.ContentType.MediaType = mediaType;
                        img.TransferEncoding = TransferEncoding.Base64;
                        img.ContentType.Name = img.ContentId;
                        img.ContentLink = new Uri("cid:" + img.ContentId);
                        LinkedResource img1 = new LinkedResource(@"C:\TataReport\TCFTemplate\approveReject.Jpeg", mediaType);
                        // Make sure you set all these values!!!
                        img1.ContentId = "EmbeddedContent_2";
                        img1.ContentType.MediaType = mediaType;
                        img1.TransferEncoding = TransferEncoding.Base64;
                        img1.ContentType.Name = img.ContentId;
                        img1.ContentLink = new Uri("cid:" + img1.ContentId);
                        //LinkedResource img2 = new LinkedResource(@"C:\TataReport\TCFTemplate\reject.Jpeg", mediaType);
                        //// Make sure you set all these values!!!
                        //img2.ContentId = "EmbeddedContent_3";
                        //img2.ContentType.MediaType = mediaType;
                        //img2.TransferEncoding = TransferEncoding.Base64;
                        //img2.ContentType.Name = img.ContentId;
                        //img2.ContentLink = new Uri("cid:" + img2.ContentId);
                        htmlView.LinkedResources.Add(img);
                        htmlView.LinkedResources.Add(img1);
                        //htmlView.LinkedResources.Add(img2);
                        mail.AlternateViews.Add(plainView);
                        mail.AlternateViews.Add(htmlView);
                    }
                    if (image == 2)
                    {
                        AlternateView htmlView = AlternateView.CreateAlternateViewFromString(message, Encoding.UTF8, MediaTypeNames.Text.Html);
                        // Create a plain text message for client that don't support HTML
                        AlternateView plainView = AlternateView.CreateAlternateViewFromString(Regex.Replace(message, "<[^>]+?>", string.Empty), Encoding.UTF8, MediaTypeNames.Text.Plain);
                        string mediaType = MediaTypeNames.Image.Jpeg;
                        LinkedResource img = new LinkedResource(@"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg", mediaType);
                        // Make sure you set all these values!!!
                        img.ContentId = "EmbeddedContent_1";
                        img.ContentType.MediaType = mediaType;
                        img.TransferEncoding = TransferEncoding.Base64;
                        img.ContentType.Name = img.ContentId;
                        img.ContentLink = new Uri("cid:" + img.ContentId);
                    }

                    //string hostName = configuration.GetSection("SMTPConn").GetSection("Host").Value;
                    //int port = Convert.ToInt32(configuration.GetSection("SMTPConn").GetSection("Port").Value);
                    //bool enableSsl = Convert.ToBoolean(configuration.GetSection("SMTPConn").GetSection("EnableSsl").Value);
                    //bool useDefaultCredentials = Convert.ToBoolean(configuration.GetSection("SMTPConn").GetSection("UseDefaultCredentials").Value);
                    //string emailId = configuration.GetSection("SMTPConn").GetSection("EmailId").Value;
                    //string password = configuration.GetSection("SMTPConn").GetSection("Password").Value;

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = hostName;
                    smtp.Port = port;
                    smtp.EnableSsl = enableSsl;
                    smtp.UseDefaultCredentials = useDefaultCredentials;
                    if (connectType == "")
                    {
                        smtp.Credentials = new System.Net.NetworkCredential(emailId, password);
                    }
                    else
                    {
                        smtp.Credentials = new System.Net.NetworkCredential(emailId, password, connectType);
                    }
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.Send(mail);


                    ret = true;
                }
            }
            catch (Exception ex)
            {
                ret = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        //Login details for Approval
        public CommonResponse1 LoginDetails(LoginInfo data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.Tblusers.Where(m => m.UserName == data.userName && m.Password == data.password).Select(m => new { m.UserId, m.PrimaryRole, m.SecondaryRole, m.UserName }).FirstOrDefault();

                if (check != null)
                {
                    var checkRole = db.Tblroles.Where(m => m.RoleId == 10).FirstOrDefault();

                    if (check.PrimaryRole == checkRole.RoleId)
                    {
                        var operatorMailId = db.Tbloperatordetails.Where(m => m.EmployeeId == check.UserName).Select(m => m.OperatorMailId).FirstOrDefault();

                        if (operatorMailId != null)
                        {
                            var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == data.machineId).Select(m => new { m.CellId, m.PlantId, m.ShopId }).FirstOrDefault();

                            string[] ids = data.tempModeIds.Split(',');
                            List<int> intArry = ids.ToList().Select(int.Parse).ToList();
                            var check1 = db.Tblunasignedwo.Where(m => intArry.Contains(m.Uawoid) && m.SendApprove == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check1.Count > 0)
                            {
                                var toMail1 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 2).Select(m => m.FirstApproverToList).FirstOrDefault();

                                if (toMail1 == operatorMailId)
                                {
                                    obj.isStatus = true;
                                    obj.response = "Valid Operator";
                                }
                                else
                                {
                                    obj.isStatus = false;
                                    obj.response = "InValid Operator";
                                }
                            }

                            var check2 = db.Tblunasignedwo.Where(m => intArry.Contains(m.Uawoid) && m.ApprovalLevel == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check2.Count > 0)
                            {
                                var toMail2 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 2).Select(m => m.SecondApproverToList).FirstOrDefault();
                                //var operatorMailId = db.Tbloperatordetails.Where(m => m.EmployeeId == checkOpDet.EmployeeId).Select(m => m.OperatorMailId).FirstOrDefault();
                                if (toMail2 == operatorMailId)
                                {
                                    obj.isStatus = true;
                                    obj.response = "Valid Operator";
                                }
                                else
                                {
                                    obj.isStatus = false;
                                    obj.response = "InValid Operator";
                                }
                            }
                        }
                    }
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "InValid User";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        public CommonResponse Edit(int uawoid)
        {
            CommonResponse comobj = new CommonResponse();
            try
            {
                var hmidet = db.Tblunasignedwo.Where(m => m.Uawoid == uawoid).FirstOrDefault();
                if (hmidet != null)
                {
                    hmidet.Status = null;
                    hmidet.Isworkinprogress = null;
                    db.SaveChanges();
                    comobj.isTure = true;
                    comobj.response = "Item Updated Successfully";
                }
            }
            catch (Exception ex)
            {

            }
            return comobj;
        }

        public CommonResponse GenericWorkOrder(GenericWO genericWO)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int switcher = 0;
                var check = (from wf in db.Tblunasignedwo
                             where wf.Uawoid == genericWO.uaWOId
                             select new
                             {
                                 Uawoid = wf.Uawoid,
                                 Starttime = wf.Starttime,
                                 Endtime = wf.Endtime
                             }).FirstOrDefault();
                if (check != null)
                {
                    var dbCheck = db.Tblmode.Where(m => m.StartTime >= check.Starttime && m.EndTime <= check.Endtime).ToList().Skip(46);
                    foreach (var item in dbCheck)
                    {
                        if (item.StartTime == check.Starttime && item.EndTime == check.Endtime)
                        {
                            switcher = 1;
                        }
                        else if (check.Starttime > item.StartTime && check.Endtime < item.EndTime && check.Endtime > item.StartTime && check.Starttime < item.EndTime)
                        {
                            switcher = 2;
                        }
                        else if (item.StartTime > check.Starttime && check.Endtime < item.EndTime && item.StartTime > check.Endtime && item.EndTime > check.Starttime)
                        {
                            switcher = 3;
                        }
                        else if (item.StartTime < check.Starttime && item.EndTime > check.Starttime && check.Endtime > item.StartTime && check.Endtime > item.EndTime)
                        {
                            switcher = 4;
                        }
                        else if (item.StartTime == check.Starttime && item.EndTime > check.Endtime && item.StartTime < check.Endtime && item.EndTime > check.Starttime)
                        {
                            switcher = 5;
                        }
                        else if (item.StartTime == check.Starttime && item.EndTime > check.Starttime && item.EndTime < check.Endtime && item.StartTime < check.Endtime)
                        {
                            switcher = 6;
                        }
                        else if (item.EndTime == check.Endtime && item.StartTime < check.Starttime && item.StartTime < check.Endtime && item.EndTime > check.Starttime)
                        {
                            switcher = 7;
                        }
                        else if (item.EndTime == check.Endtime && item.StartTime < check.Starttime && check.Starttime > item.EndTime && check.Starttime > item.StartTime)
                        {
                            switcher = 8;
                        }
                        else if (item.StartTime > check.Starttime && check.Starttime < item.EndTime && item.EndTime > check.Starttime && check.Endtime > item.EndTime)
                        {
                            switcher = 9;
                        }
                        switch (switcher)
                        {
                            case 1:
                                break;
                            case 2:

                                AddModes addModes = new AddModes();
                                addModes.machineId = item.MachineId;
                                addModes.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes.endDateTime = Convert.ToString(item.StartTime);
                                addModes.Uawoid = check.Uawoid;
                                AddNewModes(addModes);

                                addModes.machineId = item.MachineId;
                                addModes.startDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes.endDateTime = Convert.ToString(item.EndTime);
                                addModes.Uawoid = check.Uawoid;
                                AddNewModes(addModes);

                                UpdateModes updateModes = new UpdateModes();
                                updateModes.modeId = item.ModeId;
                                updateModes.startDateTime = Convert.ToString(item.StartTime);
                                updateModes.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes.Uawoid = check.Uawoid;
                                UpdateMode(updateModes);

                                break;
                            case 3:

                                AddModes addModes1 = new AddModes();
                                addModes1.machineId = item.MachineId;
                                addModes1.startDateTime = Convert.ToString(item.StartTime);
                                addModes1.endDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes1.Uawoid = check.Uawoid;
                                AddNewModes(addModes1);

                                addModes1.machineId = item.MachineId;
                                addModes1.startDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes1.endDateTime = Convert.ToString(item.EndTime);
                                addModes1.Uawoid = check.Uawoid;
                                AddNewModes(addModes1);

                                UpdateModes updateModes1 = new UpdateModes();
                                updateModes1.modeId = item.ModeId;
                                updateModes1.startDateTime = Convert.ToString(item.StartTime);
                                updateModes1.endDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes1.Uawoid = check.Uawoid;
                                UpdateMode(updateModes1);

                                break;
                            case 4:

                                AddModes addModes2 = new AddModes();
                                addModes2.machineId = item.MachineId;
                                addModes2.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes2.endDateTime = Convert.ToString(item.EndTime);
                                addModes2.Uawoid = check.Uawoid;
                                AddNewModes(addModes2);

                                addModes2.machineId = item.MachineId;
                                addModes2.startDateTime = Convert.ToString(item.EndTime);
                                addModes2.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes2.Uawoid = check.Uawoid;
                                AddNewModes(addModes2);

                                UpdateModes updateModes2 = new UpdateModes();
                                updateModes2.modeId = item.ModeId;
                                updateModes2.startDateTime = Convert.ToString(item.StartTime);
                                updateModes2.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes2.Uawoid = check.Uawoid;
                                UpdateMode(updateModes2);

                                break;

                            case 5:

                                AddModes addModes3 = new AddModes();
                                addModes3.machineId = item.MachineId;
                                addModes3.startDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes3.endDateTime = Convert.ToString(item.EndTime);
                                addModes3.Uawoid = check.Uawoid;
                                AddNewModes(addModes3);

                                UpdateModes updateModes3 = new UpdateModes();
                                updateModes3.modeId = item.ModeId;
                                updateModes3.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes3.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes3.Uawoid = check.Uawoid;
                                UpdateMode(updateModes3);
                                break;

                            case 6:

                                AddModes addModes4 = new AddModes();
                                addModes4.machineId = item.MachineId;
                                addModes4.startDateTime = Convert.ToString(item.EndTime);
                                addModes4.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes4.Uawoid = check.Uawoid;
                                AddNewModes(addModes4);

                                UpdateModes updateModes4 = new UpdateModes();
                                updateModes4.modeId = item.ModeId;
                                updateModes4.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes4.endDateTime = Convert.ToString(item.EndTime);
                                updateModes4.Uawoid = check.Uawoid;
                                UpdateMode(updateModes4);
                                break;

                            case 7:

                                AddModes addModes5 = new AddModes();
                                addModes5.machineId = item.MachineId;
                                addModes5.startDateTime = Convert.ToString(item.StartTime);
                                addModes5.endDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes5.Uawoid = check.Uawoid;
                                AddNewModes(addModes5);

                                UpdateModes updateModes5 = new UpdateModes();
                                updateModes5.modeId = item.ModeId;
                                updateModes5.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes5.endDateTime = Convert.ToString(item.EndTime);
                                updateModes5.Uawoid = check.Uawoid;
                                UpdateMode(updateModes5);

                                break;

                            case 8:

                                AddModes addModes6 = new AddModes();
                                addModes6.machineId = item.MachineId;
                                addModes6.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes6.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes6.Uawoid = check.Uawoid;
                                AddNewModes(addModes6);

                                UpdateModes updateModes6 = new UpdateModes();
                                updateModes6.modeId = item.ModeId;
                                updateModes6.startDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                updateModes6.endDateTime = Convert.ToString(item.EndTime);
                                updateModes6.Uawoid = check.Uawoid;
                                UpdateMode(updateModes6);
                                break;

                            case 9:

                                AddModes addModes7 = new AddModes();
                                addModes7.machineId = item.MachineId;
                                addModes7.startDateTime = check.Starttime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes7.endDateTime = Convert.ToString(item.StartTime);
                                addModes7.Uawoid = check.Uawoid;
                                AddNewModes(addModes7);

                                addModes7.machineId = item.MachineId;
                                addModes7.startDateTime = Convert.ToString(item.EndTime);
                                addModes7.endDateTime = check.Endtime.ToString("yyyy-MM-dd HH:mm:ss");
                                addModes7.Uawoid = check.Uawoid;
                                AddNewModes(addModes7);

                                UpdateModes updateModes7 = new UpdateModes();
                                updateModes7.modeId = item.ModeId;
                                updateModes7.startDateTime = Convert.ToString(item.StartTime);
                                updateModes7.endDateTime = Convert.ToString(item.EndTime);
                                updateModes7.Uawoid = check.Uawoid;
                                UpdateMode(updateModes7);

                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Add New Modes to LiveModeDb and LiveModeHistory table
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool AddNewModes(AddModes data)
        {
            CommonResponse obj = new CommonResponse();
            DateTime startTime = Convert.ToDateTime(data.startDateTime);
            DateTime endTime = Convert.ToDateTime(data.endDateTime);
            string getDate = DateTime.Now.ToString("yyyy-MM-dd");
            try
            {
                Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                tbllivemodedb.Mode = "PROD";
                tbllivemodedb.MachineId = data.machineId;
                tbllivemodedb.InsertedOn = DateTime.Now;
                tbllivemodedb.InsertedBy = 1;
                tbllivemodedb.IsCompleted = 1;
                tbllivemodedb.IsDeleted = 0;
                tbllivemodedb.CorrectedDate = getDate;
                tbllivemodedb.StartTime = startTime;
                tbllivemodedb.EndTime = endTime;
                tbllivemodedb.ColorCode = "GREEN";
                TimeSpan diff = endTime - startTime;
                double seconds = diff.TotalSeconds;
                tbllivemodedb.DurationInSec = Convert.ToInt32(seconds);
                db.Tbllivemodedb.Add(tbllivemodedb);
                db.SaveChanges();
                obj.isTure = true;
                obj.response = "Mode Added Successfully";

                var check = db.Tblmode.Where(m => m.ModeId == tbllivemodedb.ModeId).FirstOrDefault();
                if (check == null)
                {
                    Tblmode tblmode = new Tblmode();
                    tblmode.ModeId = tbllivemodedb.ModeId;
                    tblmode.Mode = tbllivemodedb.Mode;
                    tblmode.MachineId = tbllivemodedb.MachineId;
                    tblmode.InsertedOn = DateTime.Now;
                    tblmode.InsertedBy = 1;
                    tblmode.IsCompleted = 1;
                    tblmode.IsDeleted = 0;
                    tblmode.CorrectedDate = getDate;
                    tblmode.StartTime = startTime;
                    tblmode.EndTime = endTime;
                    tblmode.ColorCode = tbllivemodedb.ColorCode;
                    TimeSpan diff1 = endTime - startTime;
                    double seconds1 = diff.TotalSeconds;
                    tblmode.DurationInSec = Convert.ToInt32(seconds1);
                    db.Tblmode.Add(tblmode);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Mode Added Successfully";


                    var check1 = db.TblliveModeDbHis.Where(m => m.ModeId == tblmode.ModeId).FirstOrDefault();
                    if (check == null)
                    {
                        TblliveModeDbHis tblliveModeDbHis = new TblliveModeDbHis();
                        tblliveModeDbHis.ModeId = tblmode.ModeId;
                        tblliveModeDbHis.MachineId = tblmode.MachineId;
                        tblliveModeDbHis.Mode = tblmode.Mode;
                        tblliveModeDbHis.InsertedOn = DateTime.Now;
                        tblliveModeDbHis.InsertedBy = tblmode.InsertedBy;
                        tblliveModeDbHis.CorrectedDate = getDate;
                        tblliveModeDbHis.IsDeleted = tblmode.IsDeleted;
                        tblliveModeDbHis.StartTime = tblmode.StartTime;
                        tblliveModeDbHis.EndTime = tblmode.EndTime;
                        tblliveModeDbHis.ColorCode = tblmode.ColorCode;
                        tblliveModeDbHis.IsCompleted = tblmode.IsCompleted;
                        tblliveModeDbHis.DurationInSec = tblmode.DurationInSec;
                        tblliveModeDbHis.Uawoid = data.Uawoid;
                        db.TblliveModeDbHis.Add(tblliveModeDbHis);
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Mode Added Successfully";
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return true;
        }

        /// <summary>
        /// Udate mode in livemodedb and insert new mode in livemodedb history table
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public bool UpdateMode(UpdateModes data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                DateTime startTime = Convert.ToDateTime(data.startDateTime);
                DateTime endTime = Convert.ToDateTime(data.endDateTime);
                var check = db.Tblmode.Where(m => m.ModeId == data.modeId).FirstOrDefault();
                if (check != null)
                {
                    check.StartTime = startTime;
                    check.EndTime = endTime;
                    TimeSpan diff = endTime - startTime;
                    double seconds = diff.TotalSeconds;
                    check.DurationInSec = Convert.ToInt32(seconds);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Mode Updated Successfully";

                }

                var check1 = db.Tbllivemodedb.Where(m => m.ModeId == data.modeId).FirstOrDefault();
                if (check1 != null)
                {
                    check.StartTime = startTime;
                    check.EndTime = endTime;
                    TimeSpan diff = startTime - endTime;
                    double seconds = diff.TotalSeconds;
                    check.DurationInSec = Convert.ToInt32(seconds);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Mode Updated Successfully";
                }

                var checkDb = db.TblliveModeDbHis.Where(m => m.ModeId == check.ModeId).FirstOrDefault();
                string getDate = DateTime.Now.ToString("yyyy-MM-dd");
                if (checkDb == null)
                {
                    TblliveModeDbHis tblliveModeDbHis = new TblliveModeDbHis();
                    tblliveModeDbHis.ModeId = check.ModeId;
                    tblliveModeDbHis.MachineId = check.MachineId;
                    tblliveModeDbHis.Mode = check.Mode;
                    tblliveModeDbHis.InsertedOn = DateTime.Now;
                    tblliveModeDbHis.InsertedBy = check.InsertedBy;
                    tblliveModeDbHis.CorrectedDate = getDate;
                    tblliveModeDbHis.IsDeleted = check.IsDeleted;
                    tblliveModeDbHis.StartTime = check.StartTime;
                    tblliveModeDbHis.EndTime = check.EndTime;
                    tblliveModeDbHis.ColorCode = check.ColorCode;
                    tblliveModeDbHis.IsCompleted = check.IsCompleted;
                    tblliveModeDbHis.DurationInSec = check.DurationInSec;
                    tblliveModeDbHis.Uawoid = data.Uawoid;
                    db.TblliveModeDbHis.Add(tblliveModeDbHis);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Mode Added Successfully";
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return true;
        }

        public CommonResponse1 UnassignedWOSplitDurationDetails(List<UnassignedWOStartEndDateTime> data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                int i = 0;
                List<UnassignedWOStartEndDateTime> modeStartEndDateTimeList = new List<UnassignedWOStartEndDateTime>();
                foreach (var item in data)
                {
                    i++;
                    modeStartEndDateTimeList.Add(item);
                    if (data.Count == i)
                    {
                        var dbCheck = db.Tblunasignedwo.Where(m => m.Uawoid == item.UnassignedWoId).FirstOrDefault();

                        DateTime EndDateTime = Convert.ToDateTime(dbCheck.Starttime);
                        DateTime StartDateTime = Convert.ToDateTime(dbCheck.Endtime);

                        string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids = dt.Split();
                        string endDate = ids[0];
                        string endTime = ids[1];

                        string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids1 = dt1.Split();
                        string startDate = ids1[0];
                        string startTime = ids1[1];

                        string endTimeLast = ids[0] + " " + item.endTime;
                        DateTime endT = Convert.ToDateTime(endTimeLast);
                        if (endT < EndDateTime && endT > StartDateTime)
                        {
                            #region Update old row
                            dbCheck.Endtime = endT;
                            db.SaveChanges();
                            #endregion

                            #region add new row in temp mode table which we are going to insert in mode table
                            Tblunasignedwo tblunasignedwo = new Tblunasignedwo();
                            tblunasignedwo.Isdeleted = 0;
                            tblunasignedwo.Correcteddate = dbCheck.Correcteddate;
                            tblunasignedwo.Starttime = endT;
                            tblunasignedwo.Endtime = EndDateTime;
                            tblunasignedwo.Createdon = Convert.ToString(DateTime.Now);
                            tblunasignedwo.Createdby = 1;
                            tblunasignedwo.WorkOrderNo = dbCheck.WorkOrderNo;
                            tblunasignedwo.PartNo = dbCheck.PartNo;
                            tblunasignedwo.OprationNo = dbCheck.OprationNo;
                            tblunasignedwo.Project = dbCheck.Project;
                            tblunasignedwo.ProdFai = dbCheck.ProdFai;
                            tblunasignedwo.WorkOrderQty = dbCheck.WorkOrderQty;
                            tblunasignedwo.ProcessedQty = dbCheck.ProcessedQty;
                            tblunasignedwo.DeliveryQty = dbCheck.DeliveryQty;
                            tblunasignedwo.SendApprove = dbCheck.SendApprove;
                            tblunasignedwo.Acceptreject = dbCheck.Acceptreject;
                            tblunasignedwo.Isupdate = dbCheck.Isupdate;
                            tblunasignedwo.ModifiedOn = dbCheck.ModifiedOn;
                            tblunasignedwo.ModifiedBy = dbCheck.ModifiedBy;
                            tblunasignedwo.Shiftid = dbCheck.Shiftid;
                            tblunasignedwo.Operatorid = dbCheck.Operatorid;
                            tblunasignedwo.Machineid = dbCheck.Machineid;
                            tblunasignedwo.JobFinish = dbCheck.JobFinish;
                            tblunasignedwo.PartialFinish = dbCheck.PartialFinish;
                            tblunasignedwo.GenericWo = dbCheck.GenericWo;
                            tblunasignedwo.ReWork = dbCheck.ReWork;
                            tblunasignedwo.IsSplit = dbCheck.IsSplit;
                            tblunasignedwo.Loperatorid = dbCheck.Loperatorid;
                            tblunasignedwo.Status = dbCheck.Status;
                            tblunasignedwo.Isworkinprogress = dbCheck.Isworkinprogress;
                            tblunasignedwo.Isworkorder = dbCheck.Isworkorder;
                            tblunasignedwo.Pestarttime = dbCheck.Pestarttime;
                            tblunasignedwo.Ddlworkcenter = dbCheck.Ddlworkcenter;
                            tblunasignedwo.Holdcodeid = dbCheck.Holdcodeid;
                            tblunasignedwo.Holdcodereason = dbCheck.Holdcodereason;
                            tblunasignedwo.GenericCodeid = dbCheck.GenericCodeid;
                            tblunasignedwo.GenericCodereason = dbCheck.GenericCodereason;
                            tblunasignedwo.ApprovalLevel = dbCheck.ApprovalLevel;
                            tblunasignedwo.AcceptReject1 = dbCheck.AcceptReject1;
                            tblunasignedwo.RejectReason = dbCheck.RejectReason;
                            tblunasignedwo.RejectReason1 = dbCheck.RejectReason1;
                            tblunasignedwo.UpdateLevel = dbCheck.UpdateLevel;
                            db.Tblunasignedwo.Add(tblunasignedwo);
                            db.SaveChanges();
                            #endregion

                            #region Response assiging
                            UnassignedWOStartEndDateTime modeStartEndDateTime = new UnassignedWOStartEndDateTime();
                            modeStartEndDateTime.UnassignedWoId = tblunasignedwo.Uawoid;
                            modeStartEndDateTime.startDate = startDate;
                            modeStartEndDateTime.startTime = item.endTime;
                            modeStartEndDateTime.endDate = endDate;
                            modeStartEndDateTime.endTime = endTime;
                            modeStartEndDateTimeList.Add(modeStartEndDateTime);
                            obj.isStatus = true;
                            obj.response = modeStartEndDateTimeList;
                            #endregion
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.response = "Please Enter the EndDateTime with in" + EndDateTime;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        //when click on update button update the endtime
        public CommonResponse UpdateTime(CompareDuration data)
        {
            CommonResponse comobj = new CommonResponse();
            try
            {
                int uaWoId = data.uaWOId;
                string[] allIds = data.uaWOIdS.Split(',');
                for (int i = 0; i < allIds.Length; i++)
                {
                    int id = Convert.ToInt32(allIds[i]);
                    if (uaWoId == id)
                    {
                        var tcflossdet = db.Tblunasignedwo.Where(m => m.Uawoid == uaWoId).FirstOrDefault();
                        if (tcflossdet != null)
                        {
                            int count = allIds.Count();
                            int id1 = Convert.ToInt32(allIds[i + 1]);
                            int lastid = Convert.ToInt32(allIds[count - 1]);
                            var tcfdet = db.Tblunasignedwo.Where(m => m.Uawoid == id1).FirstOrDefault();
                            var lasttcfrecord = db.Tblunasignedwo.Where(m => m.Uawoid == lastid).FirstOrDefault();

                            DateTime lastetTime = Convert.ToDateTime(lasttcfrecord.Endtime);

                            DateTime tblfirststTime = Convert.ToDateTime(tcflossdet.Starttime);
                            DateTime tblEtTime = Convert.ToDateTime(tcflossdet.Endtime);
                            DateTime curEtTime = Convert.ToDateTime(data.endTime);
                            DateTime nextEtTime = Convert.ToDateTime(tcfdet.Endtime);
                            DateTime nextStTime = Convert.ToDateTime(tcfdet.Starttime);
                            DateTime PrevTime = Convert.ToDateTime(data.endTime).AddDays(-1);
                            DateTime NextTime = Convert.ToDateTime(data.endTime).AddDays(1);
                            if(curEtTime < tblfirststTime)
                            {
                                curEtTime = NextTime;
                            }
                            #region Previous Code
                            //if (curEtTime <= tblEtTime && curEtTime >= tblfirststTime)
                            //{
                            //    int durationInSecfirst = (int)curEtTime.Subtract(tblfirststTime).TotalSeconds;
                            //    int durationInSecSecond = (int)nextEtTime.Subtract(curEtTime).TotalSeconds;
                            //    if (durationInSecfirst > 120 && durationInSecSecond > 120)
                            //    {
                            //        bool check = ValidatePrvEndTime(data.uaWOIdS, PrevTime);
                            //        if (check)
                            //        {
                            //            comobj.isTure = false;
                            //            comobj.response = tblEtTime + " This Time Already Exist";
                            //        }
                            //        else
                            //        {
                            //            tcflossdet.EndDateTime = Convert.ToDateTime(data.endTime);
                            //            db.SaveChanges();

                            //            if (tcfdet != null)
                            //            {
                            //                tcfdet.StartDateTime = Convert.ToDateTime(data.endTime);
                            //                tcfdet.LossId = tcflossdet.LossId;
                            //                db.SaveChanges();
                            //            }
                            //            SplitDurationList objSplitDurationList = new SplitDurationList();
                            //            objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                            //            comobj.isTure = true;
                            //            comobj.response = objSplitDurationList;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        comobj.isTure = false;
                            //        comobj.response = "The Duration splited must be greater than 120 Seconds";
                            //    }
                            //}
                            //else if (PrevTime <= tblEtTime && PrevTime >= tblfirststTime)
                            //{
                            //    int durationInSecfirst = (int)PrevTime.Subtract(tblfirststTime).TotalSeconds;
                            //    int durationInSecSecond = (int)nextEtTime.Subtract(PrevTime).TotalSeconds;
                            //    if (durationInSecfirst > 120 && durationInSecSecond > 120)
                            //    {
                            //        bool check = ValidatePrvEndTime(data.uaWOIdS, PrevTime);
                            //        if (check)
                            //        {
                            //            comobj.isTure = false;
                            //            comobj.response = tblEtTime + " This Time Already Exist";
                            //        }
                            //        else
                            //        {

                            //            tcflossdet.EndDateTime = PrevTime;
                            //            db.SaveChanges();

                            //            if (tcfdet != null)
                            //            {
                            //                tcfdet.StartDateTime = PrevTime;
                            //                db.SaveChanges();
                            //            }
                            //            SplitDurationList objSplitDurationList = new SplitDurationList();
                            //            objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                            //            comobj.isTure = true;
                            //            comobj.response = objSplitDurationList;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        comobj.isTure = false;
                            //        comobj.response = "The Duration splited must be greater than 120 Seconds";
                            //    }
                            //}
                            //else
                            //{
                            //    comobj.isTure = false;
                            //    comobj.response = "The Time Must Be WithIn " + tblfirststTime + "-" + tblEtTime;
                            //} 
                            #endregion 

                            if (curEtTime <= lastetTime && curEtTime >= tblfirststTime)
                            {
                                if (curEtTime == nextEtTime)
                                {
                                    //bool check = ValidatePrvEndTime(data.uaWOIdS, PrevTime);
                                    //if (check)
                                    //{
                                    //    comobj.isTure = false;
                                    //    comobj.response = tblEtTime + " This Time Already Exist";
                                    //}
                                    //else
                                    //{
                                    //    db.Tbltcflossofentry.Remove(tcfdet);
                                    //    db.SaveChanges();

                                    //    tcflossdet.EndDateTime = Convert.ToDateTime(data.endTime);
                                    //    db.SaveChanges();
                                    //    SplitDurationList objSplitDurationList = new SplitDurationList();
                                    //    objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                                    //    comobj.isTure = true;
                                    //    comobj.response = objSplitDurationList;
                                    //}
                                    comobj.isTure = false;
                                    comobj.errorMsg = "You cannot able to split this duration, because splitting duration " + curEtTime + " is equal to next record end time " + nextEtTime + "";
                                }
                                else if (curEtTime > nextEtTime)
                                {
                                    comobj.isTure = false;
                                    comobj.errorMsg = "You cannot able to split this duration, because splitting duration " + curEtTime + " is Greater than next record End time " + nextEtTime + "";
                                }
                                else
                                {
                                    bool check = ValidatePrvEndTime(data.uaWOIdS, PrevTime);
                                    if (check)
                                    {
                                        comobj.isTure = false;
                                        comobj.response = tblEtTime + " This Time Already Exist";
                                    }
                                    else
                                    {
                                        tcflossdet.Endtime = curEtTime;
                                        db.SaveChanges();

                                        if (tcfdet != null)
                                        {
                                            tcfdet.Starttime = curEtTime;
                                            db.SaveChanges();
                                        }
                                        SplitDurationList objSplitDurationList = new SplitDurationList();
                                        objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                                        comobj.isTure = true;
                                        comobj.response = objSplitDurationList;
                                    }
                                }

                            }
                            else if (PrevTime <= lastetTime && PrevTime >= tblfirststTime)
                            {
                                if (curEtTime == nextEtTime)
                                {
                                    //bool check = ValidatePrvEndTime(data.uaWOIdS, PrevTime);
                                    //if (check)
                                    //{
                                    //    comobj.isTure = false;
                                    //    comobj.response = tblEtTime + " This Time Already Exist";
                                    //}
                                    //else
                                    //{
                                    //    db.Tbltcflossofentry.Remove(tcfdet);
                                    //    db.SaveChanges();

                                    //    tcflossdet.EndDateTime = PrevTime;
                                    //    db.SaveChanges();
                                    //    SplitDurationList objSplitDurationList = new SplitDurationList();
                                    //    objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                                    //    comobj.isTure = true;
                                    //    comobj.response = objSplitDurationList;
                                    //}
                                    comobj.isTure = false;
                                    comobj.errorMsg = "You cannot able to split this duration, because splitting duration " + curEtTime + " is equal to next record end time " + nextEtTime + "";
                                }
                                else
                                {
                                    tcflossdet.Endtime = PrevTime;
                                    db.SaveChanges();

                                    if (tcfdet != null)
                                    {
                                        tcfdet.Starttime = PrevTime;
                                        db.SaveChanges();
                                    }
                                    SplitDurationList objSplitDurationList = new SplitDurationList();
                                    objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                                    comobj.isTure = true;
                                    comobj.response = objSplitDurationList;
                                }
                            }
                            else
                            {
                                comobj.isTure = false;
                                comobj.response = "The Time Must Be WithIn " + tblfirststTime + "-" + lastetTime;
                            }
                        }
                        else
                        {
                            comobj.isTure = false;
                            comobj.response = ResourceResponse.NoItemsFound;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                comobj.isTure = false;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comobj;
        }

        #region Report claculation and updation
        //#region start Wo
        ////Output: In Seconds

        //public async Task<bool> CalWODataForYesterday(DateTime? StartDate, DateTime? EndDate)
        //{
        //    bool result = false;
        //    DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);
        //    //fromdate = Convert.ToDateTime(DateTime.Now.ToString("2018-05-01"));
        //    //todate = Convert.ToDateTime(DateTime.Now.ToString("2018-10-31"));
        //    if (StartDate != null && EndDate != null)
        //    {
        //        fromdate = Convert.ToDateTime(StartDate);
        //        todate = Convert.ToDateTime(EndDate);
        //    }

        //    DateTime UsedDateForExcel = Convert.ToDateTime(fromdate.ToString("yyyy-MM-dd"));
        //    double TotalDay = todate.Subtract(fromdate).TotalDays;

        //    #region
        //    for (int i = 0; i < TotalDay + 1; i++)
        //    {
        //        // 2017-03-08 
        //        string CorrectedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
        //        //Normal WorkCenter
        //        var machineData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 0).ToList();
        //        foreach (var macrow in machineData)
        //        {
        //            int MachineID = macrow.MachineId;
        //            //WorkOrder Data
        //            try
        //            {
        //                ////For Testing Just Losses
        //                //    int a = 0;
        //                //if (a == 1)
        //                //{
        //                #region
        //                var WODataPresent = db.Tblworeport.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).ToList();
        //                if (WODataPresent.Count == 0)
        //                {
        //                    var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
        //                    foreach (var hmirow in HMIData)
        //                    {
        //                        //Constants from table

        //                        int hmiid = hmirow.Hmiid;
        //                        string OperatorName = hmirow.OperatorDet;
        //                        string shift = hmirow.Shift;
        //                        string hmiCorretedDate = hmirow.CorrectedDate;
        //                        string type = hmirow.ProdFai;
        //                        string program = hmirow.Project;
        //                        int isHold = 0;
        //                        isHold = hmirow.IsHold;
        //                        DateTime StartTime = Convert.ToDateTime(hmirow.Date);
        //                        DateTime EndTime = Convert.ToDateTime(hmirow.Time);
        //                        //Values from Calculation
        //                        double cuttingTime = 0, settingTime = 0, selfInspection = 0, idle = 0, breakdown = 0, MinorLoss = 0, SummationSCTvsPP = 0;
        //                        double Blue = 0, ScrapQtyTime = 0, ReworkTime = 0;

        //                        cuttingTime = await GetGreen(CorrectedDate, StartTime, EndTime, MachineID);
        //                        cuttingTime = Math.Round(cuttingTime / 60, 2);
        //                        settingTime = await GetSettingTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        settingTime = Math.Round(settingTime / 60, 2);
        //                        selfInspection = await GetSelfInsepectionForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        selfInspection = Math.Round(selfInspection / 60, 2);
        //                        double TotalLosses = await GetAllLossesTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        TotalLosses = Math.Round(TotalLosses / 60, 2);
        //                        idle = TotalLosses;
        //                        breakdown = await GetDownTimeBreakdownForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        breakdown = Math.Round(breakdown / 60, 2);
        //                        MinorLoss = await GetMinorLossForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        MinorLoss = Math.Round(MinorLoss / 60, 2);

        //                        Blue = await GetBlue(CorrectedDate, StartTime, EndTime, MachineID);
        //                        Blue = Math.Round(Blue / 60, 2); bool isRework = false;
        //                        isRework = hmirow.IsWorkOrder == 0 ? false : true;
        //                        if (isRework)
        //                        {
        //                            ReworkTime = cuttingTime;
        //                        }

        //                        int isSingleWo = 0;
        //                        isSingleWo = hmirow.IsMultiWo;

        //                        if (isSingleWo == 0)
        //                        {
        //                            #region singleWO
        //                            string SplitWO = hmirow.SplitWo;

        //                            try
        //                            {
        //                                string PartNo = hmirow.PartNo;
        //                                string WONo = hmirow.WorkOrderNo;
        //                                string OpNo = hmirow.OperationNo;


        //                                int targetQty = Convert.ToInt32(hmirow.TargetQty);
        //                                int deliveredQty = Convert.ToInt32(hmirow.DeliveredQty);
        //                                int rejectedQty = Convert.ToInt32(hmirow.RejQty);
        //                                if (rejectedQty > 0)
        //                                {
        //                                    ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
        //                                }

        //                                int IsPF = 0;
        //                                if (hmirow.IsWorkInProgress == 1)
        //                                {
        //                                    IsPF = 1;
        //                                }

        //                                //Constants From DB
        //                                double stdCuttingTime = 0, stdMRWeight = 0;
        //                                var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
        //                                if (StdWeightTime != null)
        //                                {
        //                                    string stdCuttingTimeString = null, stdMRWeightString = null;
        //                                    string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
        //                                    stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
        //                                    stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
        //                                    stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
        //                                    stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

        //                                    double.TryParse(stdCuttingTimeString, out stdCuttingTime);
        //                                    double.TryParse(stdMRWeightString, out stdMRWeight);

        //                                    if (stdCuttingTimeUnitString == "Hrs")
        //                                    {
        //                                        stdCuttingTime = stdCuttingTime * 60;
        //                                    }
        //                                    else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
        //                                    {
        //                                        stdCuttingTime = stdCuttingTime / 60;
        //                                    }

        //                                    SummationSCTvsPP = stdCuttingTime * deliveredQty;



        //                                    // no need of else its already in minutes
        //                                }

        //                                double totalNCCuttingTime = deliveredQty * stdCuttingTime;
        //                                //??
        //                                string MRReason = null;

        //                                double WOEfficiency = 0;
        //                                if (cuttingTime != 0)
        //                                {
        //                                    WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2) * 100;
        //                                    //WOEfficiency = Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100;
        //                                }
        //                                //Now insert into table
        //                                using (i_facility_talContext db = new i_facility_talContext())
        //                                {
        //                                    Tblworeport objwo = new Tblworeport();
        //                                    objwo.MachineId = MachineID;
        //                                    objwo.Hmiid = hmiid;
        //                                    objwo.OperatorName = OperatorName;
        //                                    objwo.Shift = shift;
        //                                    objwo.CorrectedDate = hmiCorretedDate;
        //                                    objwo.PartNo = PartNo;
        //                                    objwo.WorkOrderNo = WONo;
        //                                    objwo.OpNo = OpNo;
        //                                    objwo.TargetQty = targetQty;
        //                                    objwo.DeliveredQty = deliveredQty;
        //                                    objwo.IsPf = IsPF;
        //                                    objwo.IsHold = isHold;
        //                                    objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
        //                                    objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
        //                                    objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
        //                                    objwo.Idle = (decimal)Math.Round(idle, 2);
        //                                    objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
        //                                    objwo.Type = type;
        //                                    objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
        //                                    objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
        //                                    objwo.Woefficiency = (decimal)WOEfficiency;
        //                                    objwo.RejectedQty = rejectedQty;
        //                                    objwo.Program = program;
        //                                    objwo.Mrweight = (decimal)stdMRWeight;
        //                                    objwo.InsertedOn = DateTime.Now;
        //                                    objwo.IsMultiWo = isSingleWo;
        //                                    objwo.Blue = (decimal)Math.Round(MinorLoss, 2);
        //                                    objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
        //                                    objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
        //                                    objwo.SummationOfSctvsPp = (decimal)Math.Round(SummationSCTvsPP, 2);
        //                                    objwo.StartTime = StartTime;
        //                                    objwo.EndTime = EndTime;
        //                                    db.Tblworeport.Add(objwo);
        //                                    db.SaveChanges();


        //                                    //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
        //                                    //        "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,Program,MRWeight,InsertedOn,IsMultiWO,MinorLoss,SplitWO,Blue,ScrapQtyTime,ReWorkTime,SummationOfSCTvsPP,StartTime,EndTime)"
        //                                    //        + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "','"
        //                                    //        + PartNo + "','" + WONo + "','" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "','" + Math.Round(MinorLoss, 2) + "','" + SplitWO + "','" + Math.Round(Blue, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReworkTime, 2) + "','" + Math.Round(SummationSCTvsPP, 2) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");
        //                                }
        //                            }
        //                            catch (Exception eSingle)
        //                            {
        //                                result = false;
        //                            }
        //                            #endregion
        //                        }
        //                        else
        //                        {
        //                            #region MultiWO
        //                            var MultiWOData = db.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
        //                            foreach (var multirow in MultiWOData)
        //                            {
        //                                string SplitWO = multirow.SplitWo;
        //                                try
        //                                {
        //                                    string PartNo = multirow.PartNo;
        //                                    string WONo = multirow.WorkOrder;
        //                                    string OpNo = multirow.OperationNo;
        //                                    int targetQty = Convert.ToInt32(multirow.TargetQty);
        //                                    int deliveredQty = Convert.ToInt32(multirow.DeliveredQty);
        //                                    int rejectedQty = Convert.ToInt32(multirow.ScrapQty);
        //                                    if (rejectedQty > 0)
        //                                    {
        //                                        ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
        //                                    }

        //                                    int IsPF = 0;
        //                                    if (multirow.IsCompleted == 1)
        //                                    {
        //                                        IsPF = 1;
        //                                    }
        //                                    //Constants From DB
        //                                    double stdCuttingTime = 0, stdMRWeight = 0;
        //                                    var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
        //                                    if (StdWeightTime != null)
        //                                    {
        //                                        string stdCuttingTimeString = null, stdMRWeightString = null;
        //                                        string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
        //                                        stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
        //                                        stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
        //                                        stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
        //                                        stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

        //                                        double.TryParse(stdCuttingTimeString, out stdCuttingTime);
        //                                        double.TryParse(stdMRWeightString, out stdMRWeight);

        //                                        if (stdCuttingTimeUnitString == "Hrs")
        //                                        {
        //                                            stdCuttingTime = stdCuttingTime * 60;
        //                                        }
        //                                        else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
        //                                        {
        //                                            stdCuttingTime = stdCuttingTime / 60;
        //                                        }
        //                                        SummationSCTvsPP = stdCuttingTime * deliveredQty;
        //                                    }
        //                                    double totalNCCuttingTime = deliveredQty * stdCuttingTime;
        //                                    //??
        //                                    string MRReason = null;

        //                                    double WOEfficiency = 0;
        //                                    if (cuttingTime != 0)
        //                                    {
        //                                        WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2);
        //                                    }

        //                                    //Now insert into table
        //                                    using (i_facility_talContext db = new i_facility_talContext())
        //                                    {
        //                                        try
        //                                        {
        //                                            Tblworeport objwo = new Tblworeport();
        //                                            objwo.MachineId = MachineID;
        //                                            objwo.Hmiid = hmiid;
        //                                            objwo.OperatorName = OperatorName;
        //                                            objwo.Shift = shift;
        //                                            objwo.CorrectedDate = hmiCorretedDate;
        //                                            objwo.PartNo = PartNo;
        //                                            objwo.WorkOrderNo = WONo;
        //                                            objwo.OpNo = OpNo;
        //                                            objwo.TargetQty = targetQty;
        //                                            objwo.DeliveredQty = deliveredQty;
        //                                            objwo.IsPf = IsPF;
        //                                            objwo.IsHold = isHold;
        //                                            objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
        //                                            objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
        //                                            objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
        //                                            objwo.Idle = (decimal)Math.Round(idle, 2);
        //                                            objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
        //                                            objwo.Type = type;
        //                                            objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
        //                                            objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
        //                                            objwo.Woefficiency = (decimal)WOEfficiency;
        //                                            objwo.RejectedQty = rejectedQty;
        //                                            objwo.Program = program;
        //                                            objwo.Mrweight = (decimal)stdMRWeight;
        //                                            objwo.InsertedOn = DateTime.Now;
        //                                            objwo.IsMultiWo = isSingleWo;
        //                                            objwo.Blue = (decimal)Math.Round(MinorLoss, 2);
        //                                            objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
        //                                            objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
        //                                            objwo.SummationOfSctvsPp = (decimal)Math.Round(SummationSCTvsPP, 2);
        //                                            objwo.StartTime = StartTime;
        //                                            objwo.EndTime = EndTime;
        //                                            db.Tblworeport.Add(objwo);
        //                                            db.SaveChanges();

        //                                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
        //                                            //        "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,RejectedReason,Program,MRWeight,InsertedOn,IsMultiWO,MinorLoss,SplitWO,Blue,ScrapQtyTime,ReWorkTime,SummationOfSCTvsPP,StartTime,EndTime)"
        //                                            //        + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "','"
        //                                            //        + PartNo + "','" + WONo + "','" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + MRReason + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "','" + Math.Round(MinorLoss, 2) + "','" + SplitWO + "','" + Math.Round(Blue) + "','" + Math.Round(ScrapQtyTime) + "','" + Math.Round(ReworkTime) + "','" + Math.Round(SummationSCTvsPP) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");
        //                                        }
        //                                        catch (Exception eMulti)
        //                                        {
        //                                            result = false;
        //                                        }

        //                                    }
        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    result = false;
        //                                }
        //                            }
        //                            #endregion
        //                        }
        //                    }
        //                    //result = true;
        //                }
        //                #endregion
        //            }
        //            catch (Exception ex)
        //            {
        //                result = false;
        //            }
        //            //LossesData for each WorkOrder
        //            try
        //            {
        //                #region
        //                ////Testing 
        //                //MachineID = 1;
        //                //CorrectedDate = "2017-03-22";

        //                //var HMIData = db.tblhmiscreens.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
        //                var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
        //                foreach (var hmirow in HMIData)
        //                {
        //                    int hmiid = hmirow.Hmiid;
        //                    var WODataPresent = db.Tblwolossess.Where(m => m.Hmiid == hmiid).ToList();
        //                    if (WODataPresent.Count == 0)
        //                    {
        //                        DateTime StartTime = Convert.ToDateTime(hmirow.Date);
        //                        DateTime EndTime = Convert.ToDateTime(hmirow.Time);

        //                        var LossesIDs = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeId).Distinct().ToList();
        //                        foreach (var loss in LossesIDs)
        //                        {
        //                            double duration = 0;
        //                            int lossID = loss;
        //                            using (i_facility_talContext db = new i_facility_talContext())
        //                            {
        //                                var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.MessageCodeId == lossID && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime)).ToList();

        //                                foreach (var row in query2)
        //                                {
        //                                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
        //                                    {
        //                                        DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
        //                                        DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
        //                                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                                        //Get Duration Based on start & end Time.

        //                                        if (LStartDate < StartTime)
        //                                        {
        //                                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                                            IndividualDur -= StartDurationExtra;
        //                                        }
        //                                        if (LEndDate > EndTime)
        //                                        {
        //                                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                                            IndividualDur -= EndDurationExtra;
        //                                        }
        //                                        duration += IndividualDur;
        //                                    }
        //                                }
        //                            }
        //                            if (duration > 0)
        //                            {
        //                                duration = Math.Round(duration / 60, 2);
        //                                //durationList.Add(new KeyValuePair<int, double>(lossID, duration));

        //                                //Get Loss level, and hierarchical details
        //                                int losslevel = 0, level1ID = 0, level2ID = 0;
        //                                string LossName, Level1Name, Level2Name;
        //                                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == lossID).FirstOrDefault();
        //                                int level = lossdata.LossCodesLevel;
        //                                string losscodeName = null;

        //                                #region To Get LossCode Hierarchy and Push into table
        //                                if (level == 3)
        //                                {
        //                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
        //                                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2Id);
        //                                    var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
        //                                    var lossdata2 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel2ID).FirstOrDefault();
        //                                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
        //                                    Level1Name = lossdata1.LossCode;
        //                                    Level2Name = lossdata2.LossCode;
        //                                    LossName = lossdata.LossCode;

        //                                    //Now insert into table
        //                                    using (i_facility_talContext db = new i_facility_talContext())
        //                                    {
        //                                        try
        //                                        {
        //                                            Tblwolossess objwo = new Tblwolossess();
        //                                            objwo.Hmiid = hmiid;
        //                                            objwo.LossId = lossID;
        //                                            objwo.LossName = LossName;
        //                                            objwo.LossDuration = (decimal)duration;
        //                                            objwo.Level = level;
        //                                            objwo.LossCodeLevel1Id = lossLevel1ID;
        //                                            objwo.LossCodeLevel1Name = Level1Name;
        //                                            objwo.LossCodeLevel2Id = lossLevel2ID;
        //                                            objwo.LossCodeLevel2Name = Level2Name;
        //                                            objwo.InsertedOn = DateTime.Now;
        //                                            objwo.IsDeleted = 0;
        //                                            db.Tblwolossess.Add(objwo);
        //                                            db.SaveChanges();

        //                                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                            //        + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,LossCodeLevel2ID,LossCodeLevel2Name,InsertedOn,IsDeleted) "
        //                                            //        + " VALUES('" + hmiid + "','" + lossID + "','" + LossName + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
        //                                            //        + Level1Name + "','" + lossLevel2ID + "','" + Level2Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");
        //                                        }
        //                                        catch (Exception ex)
        //                                        {
        //                                            result = false;
        //                                        }
        //                                    }


        //                                }
        //                                else if (level == 2)
        //                                {
        //                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
        //                                    var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
        //                                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
        //                                    Level1Name = lossdata1.LossCode;

        //                                    //Now insert into table
        //                                    using (i_facility_talContext db = new i_facility_talContext())
        //                                    {
        //                                        try
        //                                        {
        //                                            Tblwolossess objwo = new Tblwolossess();
        //                                            objwo.Hmiid = hmiid;
        //                                            objwo.LossId = lossID;
        //                                            objwo.LossName = lossdata.LossCode;
        //                                            objwo.LossDuration = (decimal)duration;
        //                                            objwo.Level = level;
        //                                            objwo.LossCodeLevel1Id = lossLevel1ID;
        //                                            objwo.LossCodeLevel1Name = Level1Name;
        //                                            objwo.InsertedOn = DateTime.Now;
        //                                            objwo.IsDeleted = 0;
        //                                            db.Tblwolossess.Add(objwo);
        //                                            db.SaveChanges();

        //                                            //                        SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                            //+ "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,InsertedOn,IsDeleted) "
        //                                            //+ " VALUES('" + hmiid + "','" + lossID + "','" + lossdata.LossCode + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
        //                                            //+ Level1Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");
        //                                        }
        //                                        catch (Exception e)
        //                                        {
        //                                            result = false;
        //                                        }
        //                                    }

        //                                }
        //                                else if (level == 1)
        //                                {
        //                                    if (lossID == 999)
        //                                    {
        //                                        losscodeName = "NoCode Entered";
        //                                    }
        //                                    else
        //                                    {
        //                                        losscodeName = lossdata.LossCode;
        //                                    }
        //                                    //Now insert into table
        //                                    using (i_facility_talContext db = new i_facility_talContext())
        //                                    {
        //                                        try
        //                                        {
        //                                            Tblwolossess objwo = new Tblwolossess();
        //                                            objwo.Hmiid = hmiid;
        //                                            objwo.LossId = lossID;
        //                                            objwo.LossName = lossdata.LossCode;
        //                                            objwo.LossDuration = (decimal)duration;
        //                                            objwo.Level = level;
        //                                            objwo.InsertedOn = DateTime.Now;
        //                                            objwo.IsDeleted = 0;
        //                                            db.Tblwolossess.Add(objwo);
        //                                            db.SaveChanges();

        //                                            //                        SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                            //+ "(HMIID,LossID,LossName,LossDuration,Level,InsertedOn,IsDeleted) "
        //                                            //+ " VALUES('" + hmiid + "','" + lossID + "','" + losscodeName + "','" + duration + "','" + level + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0);");
        //                                        }
        //                                        catch (Exception e)
        //                                        {
        //                                            result = false;
        //                                        }
        //                                    }
        //                                }
        //                                #endregion

        //                            }
        //                        }
        //                    }

        //                }
        //                //result = true;
        //                #endregion
        //            }
        //            catch (Exception e)
        //            {

        //            }
        //        }

        //        //For Manual WorkCenters.
        //        var MWCData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 1 && m.ManualWcid.HasValue).ToList();
        //        foreach (var macrow in MWCData)
        //        {
        //            int MachineID = macrow.MachineId;
        //            try
        //            {
        //                #region
        //                var WODataPresent = db.Tblworeport.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).ToList();
        //                if (WODataPresent.Count == 0)
        //                {
        //                    var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
        //                    foreach (var hmirow in HMIData)
        //                    {
        //                        //Constants from table

        //                        int hmiid = hmirow.Hmiid;
        //                        string OperatorName = hmirow.OperatorDet;
        //                        string shift = hmirow.Shift;
        //                        string hmiCorretedDate = hmirow.CorrectedDate;
        //                        string type = hmirow.ProdFai;
        //                        string program = hmirow.Project;
        //                        int isHold = 0;
        //                        isHold = hmirow.IsHold;
        //                        string SplitWO = hmirow.SplitWo;
        //                        int HoldID = 0; string HoldReasonID = null;
        //                        try
        //                        {
        //                            HoldID = Convert.ToInt32(db.Tblmanuallossofentry.Where(m => m.Hmiid == hmiid).Select(m => m.MessageCodeId).FirstOrDefault());
        //                        }
        //                        catch (Exception e)
        //                        {

        //                        }
        //                        if (HoldID != 0)
        //                        {
        //                            HoldReasonID = HoldID.ToString();
        //                        }

        //                        DateTime StartTime = Convert.ToDateTime(hmirow.Date);
        //                        DateTime EndTime = Convert.ToDateTime(hmirow.Time);
        //                        //Values from Calculation
        //                        double cuttingTime = 0, settingTime = 0, selfInspection = 0, idle = 0, breakdown = 0;
        //                        double Blue = 0, ScrapQtyTime = 0, ReworkTime = 0;

        //                        settingTime = await GetSettingTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        settingTime = Math.Round(settingTime / 60, 2);
        //                        selfInspection = await GetSelfInsepectionForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        selfInspection = Math.Round(selfInspection / 60, 2);
        //                        double TotalLosses = await GetAllLossesTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
        //                        TotalLosses = Math.Round(TotalLosses / 60, 2);
        //                        idle = TotalLosses;
        //                        breakdown = 0;

        //                        var HMIIDData = db.Tblhmiscreen.Where(m => m.Hmiid == hmiid).FirstOrDefault();
        //                        DateTime WOStartDateTime = Convert.ToDateTime(HMIIDData.Date);
        //                        DateTime WOEndDateTime = Convert.ToDateTime(HMIIDData.Time);
        //                        double TotalWODurationIsSec = WOEndDateTime.Subtract(WOStartDateTime).TotalMinutes;

        //                        cuttingTime = TotalWODurationIsSec - settingTime - selfInspection;

        //                        int isSingleWo = 0;
        //                        isSingleWo = hmirow.IsMultiWo;

        //                        try
        //                        {
        //                            string PartNo = hmirow.PartNo;
        //                            string WONo = hmirow.WorkOrderNo;
        //                            string OpNo = hmirow.OperationNo;
        //                            int targetQty = Convert.ToInt32(hmirow.TargetQty);
        //                            int deliveredQty = Convert.ToInt32(hmirow.DeliveredQty);
        //                            int rejectedQty = Convert.ToInt32(hmirow.RejQty);
        //                            int IsPF = 0;
        //                            if (hmirow.IsWorkInProgress == 1)
        //                            {
        //                                IsPF = 1;
        //                            }

        //                            if (rejectedQty > 0)
        //                            {
        //                                ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
        //                            }

        //                            bool isRework = false;
        //                            isRework = hmirow.IsWorkOrder == 1 ? true : false;
        //                            if (isRework)
        //                            {
        //                                ReworkTime = cuttingTime;
        //                            }

        //                            //Constants From DB
        //                            double stdCuttingTime = 0, stdMRWeight = 0;
        //                            var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
        //                            if (StdWeightTime != null)
        //                            {
        //                                string stdCuttingTimeString = null, stdMRWeightString = null;
        //                                string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
        //                                stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
        //                                stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
        //                                stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
        //                                stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

        //                                double.TryParse(stdCuttingTimeString, out stdCuttingTime);
        //                                double.TryParse(stdMRWeightString, out stdMRWeight);

        //                                stdCuttingTimeUnitString = StdWeightTime.StdCuttingTimeUnit;
        //                                stdCuttingTimeUnitString = StdWeightTime.StdCuttingTimeUnit;

        //                                if (stdCuttingTimeUnitString == "Hrs")
        //                                {
        //                                    stdCuttingTime = stdCuttingTime * 60;
        //                                }
        //                                else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
        //                                {
        //                                    stdCuttingTime = stdCuttingTime / 60;
        //                                }
        //                            }
        //                            double totalNCCuttingTime = deliveredQty * stdCuttingTime;
        //                            //??
        //                            string MRReason = null;

        //                            double WOEfficiency = 0;
        //                            if (cuttingTime != 0)
        //                            {
        //                                WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2) * 100;
        //                                //WOEfficiency = Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100;
        //                            }
        //                            //Now insert into table

        //                            using (i_facility_talContext db = new i_facility_talContext())
        //                            {
        //                                try
        //                                {
        //                                    Tblworeport objwo = new Tblworeport();
        //                                    objwo.MachineId = MachineID;
        //                                    objwo.Hmiid = hmiid;
        //                                    objwo.OperatorName = OperatorName;
        //                                    objwo.Shift = shift;
        //                                    objwo.CorrectedDate = hmiCorretedDate;
        //                                    objwo.PartNo = PartNo;
        //                                    objwo.WorkOrderNo = WONo;
        //                                    objwo.OpNo = OpNo;
        //                                    objwo.TargetQty = targetQty;
        //                                    objwo.DeliveredQty = deliveredQty;
        //                                    objwo.IsPf = IsPF;
        //                                    objwo.IsHold = isHold;
        //                                    objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
        //                                    objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
        //                                    objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
        //                                    objwo.Idle = (decimal)Math.Round(idle, 2);
        //                                    objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
        //                                    objwo.Type = type;
        //                                    objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
        //                                    objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
        //                                    objwo.Woefficiency = (decimal)WOEfficiency;
        //                                    objwo.RejectedQty = rejectedQty;
        //                                    objwo.Program = program;
        //                                    objwo.Mrweight = (decimal)stdMRWeight;
        //                                    objwo.InsertedOn = DateTime.Now;
        //                                    objwo.IsMultiWo = isSingleWo;
        //                                    objwo.IsNormalWc = 1;
        //                                    objwo.HoldReason = HoldReasonID;
        //                                    objwo.SplitWo = SplitWO;
        //                                    objwo.Blue = (decimal)Math.Round(Blue, 2);
        //                                    objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
        //                                    objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
        //                                    objwo.StartTime = StartTime;
        //                                    objwo.EndTime = EndTime;
        //                                    db.Tblworeport.Add(objwo);
        //                                    db.SaveChanges();

        //                                    //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
        //                                    //    "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,Program,MRWeight,InsertedOn,IsMultiWO,IsNormalWC,HoldReason,SplitWO,Blue,ScrapQtyTime,ReWorkTime,StartTime, EndTime)"
        //                                    //    + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "',\""
        //                                    //    + PartNo + "\",\"" + WONo + "\",'" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "',1,'" + HoldReasonID + "','" + SplitWO + "','" + Math.Round(Blue, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReworkTime, 2) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");

        //                                }
        //                                catch (Exception e)
        //                                {
        //                                    result = false;
        //                                }
        //                            }
        //                        }
        //                        catch (Exception eSingle)
        //                        {
        //                            result = false;
        //                        }

        //                    }
        //                }
        //                #endregion
        //            }
        //            catch (Exception e)
        //            {
        //                result = false;
        //            }

        //            //LossesData for each WorkOrder
        //            try
        //            {
        //                #region

        //                var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
        //                foreach (var hmirow in HMIData)
        //                {
        //                    int hmiid = hmirow.Hmiid;
        //                    var WODataPresent = db.Tblwolossess.Where(m => m.Hmiid == hmiid).ToList();
        //                    if (WODataPresent.Count == 0)
        //                    {
        //                        DateTime StartTime = Convert.ToDateTime(hmirow.Date);
        //                        DateTime EndTime = Convert.ToDateTime(hmirow.Time);

        //                        var LossesIDs = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeId).Distinct().ToList();
        //                        foreach (var loss in LossesIDs)
        //                        {

        //                            double duration = 0;
        //                            int lossID = loss;
        //                            using (i_facility_talContext db = new i_facility_talContext())
        //                            {
        //                                var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.MessageCodeId == lossID && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime));


        //                                foreach (var row in query2)
        //                                {
        //                                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
        //                                    {
        //                                        DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
        //                                        DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
        //                                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                                        //Get Duration Based on start & end Time.

        //                                        if (LStartDate < StartTime)
        //                                        {
        //                                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                                            IndividualDur -= StartDurationExtra;
        //                                        }
        //                                        if (LEndDate > EndTime)
        //                                        {
        //                                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                                            IndividualDur -= EndDurationExtra;
        //                                        }
        //                                        duration += IndividualDur;
        //                                    }
        //                                }



        //                                if (duration > 0)
        //                                {
        //                                    duration = Math.Round(duration / 60, 2);
        //                                    //durationList.Add(new KeyValuePair<int, double>(lossID, duration));

        //                                    //Get Loss level, and hierarchical details
        //                                    string LossName, Level1Name, Level2Name;
        //                                    var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == lossID).FirstOrDefault();
        //                                    int level = lossdata.LossCodesLevel;
        //                                    string losscodeName = null;

        //                                    #region To Get LossCode Hierarchy and Push into table
        //                                    if (level == 3)
        //                                    {
        //                                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
        //                                        int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2Id);
        //                                        var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
        //                                        var lossdata2 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel2ID).FirstOrDefault();
        //                                        losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
        //                                        Level1Name = lossdata1.LossCode;
        //                                        Level2Name = lossdata2.LossCode;
        //                                        LossName = lossdata.LossCode;

        //                                        //Now insert into table
        //                                        using (i_facility_talContext db1 = new i_facility_talContext())
        //                                        {
        //                                            try
        //                                            {
        //                                                Tblwolossess objwo = new Tblwolossess();
        //                                                objwo.Hmiid = hmiid;
        //                                                objwo.LossId = lossID;
        //                                                objwo.LossName = LossName;
        //                                                objwo.LossDuration = (decimal)duration;
        //                                                objwo.Level = level;
        //                                                objwo.LossCodeLevel1Id = lossLevel1ID;
        //                                                objwo.LossCodeLevel1Name = Level1Name;
        //                                                objwo.LossCodeLevel2Id = lossLevel2ID;
        //                                                objwo.LossCodeLevel2Name = Level2Name;
        //                                                objwo.InsertedOn = DateTime.Now;
        //                                                objwo.IsDeleted = 0;
        //                                                db1.Tblwolossess.Add(objwo);
        //                                                db1.SaveChanges();
        //                                                //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                                //    + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,LossCodeLevel2ID,LossCodeLevel2Name,InsertedOn,IsDeleted) "
        //                                                //    + " VALUES('" + hmiid + "','" + lossID + "','" + LossName + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
        //                                                //    + Level1Name + "','" + lossLevel2ID + "','" + Level2Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");

        //                                            }
        //                                            catch (Exception e)
        //                                            {
        //                                                result = false;
        //                                            }
        //                                        }

        //                                    }
        //                                    else if (level == 2)
        //                                    {
        //                                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
        //                                        var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
        //                                        losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
        //                                        Level1Name = lossdata1.LossCode;

        //                                        //Now insert into table
        //                                        using (i_facility_talContext db1 = new i_facility_talContext())
        //                                        {
        //                                            try
        //                                            {
        //                                                Tblwolossess objwo = new Tblwolossess();
        //                                                objwo.Hmiid = hmiid;
        //                                                objwo.LossId = lossID;
        //                                                objwo.LossName = lossdata.LossCode;
        //                                                objwo.LossDuration = (decimal)duration;
        //                                                objwo.Level = level;
        //                                                objwo.LossCodeLevel1Id = lossLevel1ID;
        //                                                objwo.LossCodeLevel1Name = Level1Name;
        //                                                objwo.InsertedOn = DateTime.Now;
        //                                                objwo.IsDeleted = 0;
        //                                                db1.Tblwolossess.Add(objwo);
        //                                                db1.SaveChanges();
        //                                                //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                                //    + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,InsertedOn,IsDeleted) "
        //                                                //    + " VALUES('" + hmiid + "','" + lossID + "','" + lossdata.LossCode + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
        //                                                //    + Level1Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");

        //                                            }
        //                                            catch (Exception e)
        //                                            {
        //                                                result = false;
        //                                            }
        //                                        }

        //                                    }
        //                                    else if (level == 1)
        //                                    {
        //                                        if (lossID == 999)
        //                                        {
        //                                            losscodeName = "NoCode Entered";
        //                                        }
        //                                        else
        //                                        {
        //                                            losscodeName = lossdata.LossCode;
        //                                        }
        //                                        //Now insert into table
        //                                        using (i_facility_talContext db1 = new i_facility_talContext())
        //                                        {
        //                                            try
        //                                            {
        //                                                Tblwolossess objwo = new Tblwolossess();
        //                                                objwo.Hmiid = hmiid;
        //                                                objwo.LossId = lossID;
        //                                                objwo.LossName = lossdata.LossCode;
        //                                                objwo.LossDuration = (decimal)duration;
        //                                                objwo.Level = level;
        //                                                objwo.InsertedOn = DateTime.Now;
        //                                                objwo.IsDeleted = 0;
        //                                                db1.Tblwolossess.Add(objwo);
        //                                                db1.SaveChanges();
        //                                                //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
        //                                                //    + "(HMIID,LossID,LossName,LossDuration,Level,InsertedOn,IsDeleted) "
        //                                                //    + " VALUES('" + hmiid + "','" + lossID + "','" + losscodeName + "','" + duration + "','" + level + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0);");

        //                                            }
        //                                            catch (Exception e)
        //                                            {
        //                                                result = false;
        //                                            }
        //                                        }
        //                                    }
        //                                    #endregion
        //                                }
        //                            }
        //                        }
        //                    }

        //                    #endregion
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                result = false;
        //            }
        //        }
        //        result = true;
        //        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
        //    }
        //    #endregion

        //    return await Task.FromResult<bool>(result);

        //}

        //public async Task<double> GetSettingTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        //{
        //    double settingTime = 0;
        //    int setupid = 0;
        //    string settingString = "Setup";
        //    var setupiddata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.MessageType.Equals(settingString, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
        //    if (setupiddata != null)
        //    {
        //        setupid = setupiddata.LossCodeId;
        //    }
        //    else
        //    {
        //        return -1;
        //    }

        //    //var s = string.Join(",", products.Where(p => p.ProductType == someType).Select(p => p.ProductId.ToString()));
        //    // getting all setup's sublevels ids.
        //    var SettingIDs = db.Tbllossescodes
        //                        .Where(m => m.LossCodesLevel1Id == setupid)
        //                        .Select(m => m.LossCodeId).ToList()
        //                        .Distinct();
        //    string SettingIDsString = null;
        //    int j = 0;
        //    foreach (var row in SettingIDs)
        //    {
        //        if (j != 0)
        //        {
        //            SettingIDsString += "," + Convert.ToInt32(row);
        //        }
        //        else
        //        {
        //            SettingIDsString = Convert.ToInt32(row).ToString();
        //        }
        //        j++;
        //    }

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && SettingIDs.Contains(m.MessageCodeId) && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime)).ToList();


        //        foreach (var row in query2)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                settingTime += IndividualDur;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(settingTime);
        //}

        //public async Task<double> GetSelfInsepectionForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        //{
        //    double SelfInspectionTime = 0;
        //    int SelfInspectionid = 112;

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query3 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.MessageCodeId == SelfInspectionid && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime)).ToList();
        //        foreach (var row in query3)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                SelfInspectionTime += IndividualDur;
        //            }
        //        }
        //    }

        //    return await Task.FromResult<double>(SelfInspectionTime);
        //}

        //public async Task<double> GetAllLossesTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        //{
        //    double AllLossesTime = 0;

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query3 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime)).ToList();

        //        foreach (var row in query3)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                AllLossesTime += IndividualDur;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(AllLossesTime);
        //}

        //public async Task<double> GetDownTimeBreakdownForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        //{
        //    double BreakdownTime = 0;

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query3 = db.Tblbreakdown.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1 && m.StartTime <= StartTime && m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime) || (m.StartTime > StartTime && m.StartTime < EndTime)).ToList();

        //        foreach (var row in query3)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;
        //                // Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                BreakdownTime += IndividualDur;
        //            }

        //        }
        //    }

        //    return await Task.FromResult<double>(BreakdownTime);
        //}

        //public async Task<double> GetMinorLossForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        //{
        //    double MinorLoss = 0;

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
        //          && m.ColorCode == "Yellow" && m.DurationInSec < 120 && m.StartTime <= StartTime && m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime) || (m.StartTime > StartTime && m.StartTime < EndTime)).ToList();

        //        foreach (var row in query1)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                MinorLoss += IndividualDur;
        //            }
        //        }

        //    }

        //    return await Task.FromResult<double>(MinorLoss);
        //}

        //public void DeletePrvDaysDataFromLiveDPS()
        //{
        //    try
        //    {
        //        string CorrectedDate = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd");
        //        using (i_facility_talContext dblivedps = new i_facility_talContext())
        //        {
        //            var liveDPSData = dblivedps.Tbllivedailyprodstatus.Where(m => m.CorrectedDate == CorrectedDate).ToList();
        //            if (liveDPSData != null)
        //            {
        //                dblivedps.Tbllivedailyprodstatus.RemoveRange(liveDPSData);
        //                dblivedps.SaveChanges();
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {

        //    }
        //}

        //////Output: Seconds.
        //public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID, int HMIID)
        //{
        //    double SQT = 0;
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsData = dbhmi.Tblhmiscreen.Where(m => m.Hmiid == HMIID).FirstOrDefault();
        //        if (PartsData != null)
        //        {
        //            int scrapQty = Convert.ToInt32(PartsData.RejQty);
        //            int DeliveredQty = Convert.ToInt32(PartsData.DeliveredQty);
        //            Double WODuration = await GetGreen(UsedDateForExcel, StartTime, EndTime, MachineID);
        //            if ((scrapQty + DeliveredQty) == 0)
        //            {
        //                SQT += 0;
        //            }
        //            else
        //            {
        //                SQT += (WODuration / (scrapQty + DeliveredQty)) * scrapQty;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(SQT);
        //}

        //////Output: Seconds
        //public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID, int HMIID)
        //{
        //    double SQT = 0;
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsData = dbhmi.Tblhmiscreen.Where(m => m.Hmiid == HMIID).FirstOrDefault();
        //        if (PartsData != null)
        //        {
        //            int scrapQty = Convert.ToInt32(PartsData.RejQty);
        //            int DeliveredQty = Convert.ToInt32(PartsData.DeliveredQty);
        //            SQT = await GetGreen(UsedDateForExcel, StartTime, EndTime, MachineID);
        //        }
        //    }
        //    return await Task.FromResult<double>(SQT);
        //}

        //////Output: Minutes
        //public async Task<double> GetSummationOfSCTvsPPForWO(int HMIID)
        //{
        //    double SummationofTime = 0;
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.Hmiid == HMIID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
        //        if (PartsDataAll.Count == 0)
        //        {
        //            //return SummationofTime;
        //        }
        //        foreach (var row in PartsDataAll)
        //        {
        //            if (row.IsMultiWo == 0)
        //            {
        //                string partNo = row.PartNo;
        //                string woNo = row.WorkOrderNo;
        //                string opNo = row.OperationNo;
        //                int DeliveredQty = 0;
        //                DeliveredQty = Convert.ToInt32(row.DeliveredQty);
        //                #region InnerLogic Common for both ways(HMI or tblmultiWOselection)
        //                double stdCuttingTime = 0;
        //                var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //                if (stdcuttingTimeData != null)
        //                {
        //                    double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //                    string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //                    if (Unit == "Hrs")
        //                    {
        //                        stdCuttingTime = stdcuttingval * 60;
        //                    }
        //                    else if (Unit == "Sec") //Unit is Minutes
        //                    {
        //                        stdCuttingTime = stdcuttingval / 60;
        //                    }
        //                    else
        //                    {
        //                        stdCuttingTime = stdcuttingval;
        //                    }
        //                }
        //                #endregion
        //                SummationofTime += stdCuttingTime * DeliveredQty;
        //            }
        //            else
        //            {
        //                int hmiid = row.Hmiid;
        //                var multiWOData = dbhmi.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
        //                foreach (var rowMulti in multiWOData)
        //                {
        //                    string partNo = rowMulti.PartNo;
        //                    string opNo = rowMulti.OperationNo;
        //                    int DeliveredQty = 0;
        //                    DeliveredQty = Convert.ToInt32(rowMulti.DeliveredQty);
        //                    #region
        //                    double stdCuttingTime = 0;
        //                    var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //                    if (stdcuttingTimeData != null)
        //                    {
        //                        double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //                        string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //                        if (Unit == "Hrs")
        //                        {
        //                            stdCuttingTime = stdcuttingval * 60;
        //                        }
        //                        else if (Unit == "Sec") //Unit is Minutes
        //                        {
        //                            stdCuttingTime = stdcuttingval / 60;
        //                        }
        //                        else
        //                        {
        //                            stdCuttingTime = stdcuttingval;
        //                        }

        //                    }
        //                    #endregion
        //                    SummationofTime += stdCuttingTime * DeliveredQty;
        //                }
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(SummationofTime);
        //}

        //#endregion WO

        //#region UpdateOEETable

        //public async Task<bool> CalculateOEEForYesterday(DateTime? StartDate, DateTime? EndDate)
        //{
        //    bool result = false;
        //    //MessageBox.Show("StartTime= " + StartDate + " EndTime= " + EndDate);

        //    DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);

        //    if (StartDate != null && EndDate != null)
        //    {
        //        fromdate = Convert.ToDateTime(StartDate);
        //        todate = Convert.ToDateTime(EndDate);
        //    }
        //    //fromdate = Convert.ToDateTime(DateTime.Now.ToString("2018-05-01"));
        //    //todate = Convert.ToDateTime(DateTime.Now.ToString("2018-10-31"));

        //    //commented by V For calculating  sent date
        //    //fromdate = StartDate ?? DateTime.Now.AddDays(-1);
        //    //todate = EndDate ?? DateTime.Now.AddDays(-1);

        //    //DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);
        //    DateTime UsedDateForExcel = Convert.ToDateTime(fromdate.ToString("yyyy-MM-dd 00:00:00"));
        //    double TotalDay = todate.Subtract(fromdate).TotalDays;
        //    #region
        //    for (int i = 0; i < TotalDay + 1; i++)
        //    {
        //        //2017 - 02 - 17
        //        var machineData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 0).ToList();
        //        foreach (var macrow in machineData)
        //        {
        //            int MachineID = macrow.MachineId;

        //            try
        //            {
        //                var OEEDataPresent = db.Tbloeedashboardvariables.Where(m => m.Wcid == MachineID && m.StartDate == UsedDateForExcel).ToList();
        //                if (OEEDataPresent.Count == 0)
        //                {
        //                    double green, red, yellow, blue, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
        //                    double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
        //                    double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;

        //                    MinorLosses = await GetMinorLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                    if (MinorLosses < 0)
        //                    {
        //                        MinorLosses = 0;
        //                    }
        //                    blue = await GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                    green = await GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");

        //                    try
        //                    {
        //                        //Availability
        //                        SettingTime = await GetSettingTime(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (SettingTime < 0)
        //                        {
        //                            SettingTime = 0;
        //                        }
        //                        ROALossess = await GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROA");
        //                        if (ROALossess < 0)
        //                        {
        //                            ROALossess = 0;
        //                        }
        //                        DownTimeBreakdown = await GetDownTimeBreakdown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (DownTimeBreakdown < 0)
        //                        {
        //                            DownTimeBreakdown = 0;
        //                        }

        //                        //Performance
        //                        SummationOfSCTvsPP = await GetSummationOfSCTvsPP(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (SummationOfSCTvsPP <= 0)
        //                        {
        //                            SummationOfSCTvsPP = 0;
        //                        }

        //                        //ROPLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROP");
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        result = false;
        //                    }

        //                    //Quality
        //                    try
        //                    {
        //                        ScrapQtyTime = await GetScrapQtyTimeOfWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (ScrapQtyTime < 0)
        //                        {
        //                            ScrapQtyTime = 0;
        //                        }
        //                        ReWOTime = await GetScrapQtyTimeOfRWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (ReWOTime < 0)
        //                        {
        //                            ReWOTime = 0;
        //                        }
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        result = false;
        //                    }
        //                    //Take care when using Available Time in Calculation of OEE and Stuff.
        //                    //if (TimeType == "GodHours")
        //                    //{
        //                    //    AvailableTime = AvailableTime = 24 * 60; //24Hours to Minutes;
        //                    //}

        //                    OperatingTime = green;

        //                    //To get Top 5 Losses for this WC
        //                    string todayAsCorrectedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
        //                    DataTable DTLosses = new DataTable();
        //                    DTLosses.Columns.Add("lossCodeID", typeof(int));
        //                    DTLosses.Columns.Add("LossDuration", typeof(int));


        //                    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //                    {
        //                        var lossData = dbLoss.Tbllossofentry.Where(m => m.CorrectedDate == todayAsCorrectedDate && m.MachineId == MachineID).ToList();
        //                        foreach (var row in lossData)
        //                        {
        //                            int lossCodeID = Convert.ToInt32(row.MessageCodeId);
        //                            DateTime startDate = Convert.ToDateTime(row.StartDateTime);
        //                            DateTime endDate = Convert.ToDateTime(row.EndDateTime);
        //                            int duration = Convert.ToInt32(endDate.Subtract(startDate).TotalMinutes);

        //                            DataRow dr = DTLosses.Select("lossCodeID= '" + lossCodeID + "'").FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
        //                            if (dr != null)
        //                            {
        //                                int LossDurationPrev = Convert.ToInt32(dr["LossDuration"]); //get lossduration and update it.
        //                                dr["LossDuration"] = (LossDurationPrev + duration);
        //                            }
        //                            //}
        //                            else
        //                            {
        //                                DTLosses.Rows.Add(lossCodeID, duration);
        //                            }
        //                        }
        //                    }
        //                    DataTable DTLossesTop5 = DTLosses.Clone();
        //                    //get only the rows you want
        //                    DataRow[] results = DTLosses.Select("", "LossDuration DESC");
        //                    //populate new destination table
        //                    if (DTLosses.Rows.Count > 0)
        //                    {
        //                        int num = DTLosses.Rows.Count;
        //                        for (var iDT = 0; iDT < num; iDT++)
        //                        {
        //                            if (results[iDT] != null)
        //                            {
        //                                DTLossesTop5.ImportRow(results[iDT]);
        //                            }
        //                            else
        //                            {
        //                                DTLossesTop5.Rows.Add(0, 0);
        //                            }
        //                            if (iDT == 4)
        //                            {
        //                                break;
        //                            }
        //                        }
        //                        if (num < 5)
        //                        {
        //                            for (var iDT = num; iDT < 5; iDT++)
        //                            {
        //                                DTLossesTop5.Rows.Add(0, 0);
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        for (var iDT = 0; iDT < 5; iDT++)
        //                        {
        //                            DTLossesTop5.Rows.Add(0, 0);
        //                        }
        //                    }
        //                    ////Gather LossValues
        //                    string lossCode1, lossCode2, lossCode3, lossCode4, lossCode5 = null;
        //                    int lossCodeVal1, lossCodeVal2, lossCodeVal3, lossCodeVal4, lossCodeVal5 = 0;

        //                    lossCode1 = Convert.ToString(DTLossesTop5.Rows[0][0]);
        //                    lossCode2 = Convert.ToString(DTLossesTop5.Rows[1][0]);
        //                    lossCode3 = Convert.ToString(DTLossesTop5.Rows[2][0]);
        //                    lossCode4 = Convert.ToString(DTLossesTop5.Rows[3][0]);
        //                    lossCode5 = Convert.ToString(DTLossesTop5.Rows[4][0]);
        //                    lossCodeVal1 = Convert.ToInt32(DTLossesTop5.Rows[0][1]);
        //                    lossCodeVal2 = Convert.ToInt32(DTLossesTop5.Rows[1][1]);
        //                    lossCodeVal3 = Convert.ToInt32(DTLossesTop5.Rows[2][1]);
        //                    lossCodeVal4 = Convert.ToInt32(DTLossesTop5.Rows[3][1]);
        //                    lossCodeVal5 = Convert.ToInt32(DTLossesTop5.Rows[4][1]);

        //                    //Gather Plant, Shop, Cell for WC.

        //                    //int PlantID = 0, ShopID = 0, CellID = 0;
        //                    string PlantIDS = null, ShopIDS = null, CellIDS = null;
        //                    int value;
        //                    var WCData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == MachineID).FirstOrDefault();
        //                    string TempVal = WCData.PlantId.ToString();
        //                    if (int.TryParse(TempVal, out value))
        //                    {
        //                        PlantIDS = value.ToString();
        //                    }

        //                    TempVal = WCData.ShopId.ToString();
        //                    if (int.TryParse(TempVal, out value))
        //                    {
        //                        ShopIDS = value.ToString();
        //                    }

        //                    TempVal = WCData.CellId.ToString();
        //                    if (int.TryParse(TempVal, out value))
        //                    {
        //                        CellIDS = value.ToString();
        //                    }

        //                    // Now insert into table
        //                    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //                    {
        //                        try
        //                        {
        //                            Tbloeedashboardvariables objoee = new Tbloeedashboardvariables();
        //                            objoee.PlantId = Convert.ToInt32(PlantIDS);
        //                            objoee.ShopId = Convert.ToInt32(ShopIDS);
        //                            objoee.CellId = Convert.ToInt32(CellIDS);
        //                            objoee.Wcid = Convert.ToInt32(MachineID);
        //                            objoee.StartDate = UsedDateForExcel;
        //                            objoee.EndDate = UsedDateForExcel;
        //                            objoee.MinorLosses = Math.Round(MinorLosses / 60, 2);
        //                            objoee.Blue = Math.Round(blue / 60, 2);
        //                            objoee.Green = Math.Round(green / 60, 2);
        //                            objoee.SettingTime = Math.Round(SettingTime, 2);
        //                            objoee.Roalossess = Math.Round(ROALossess / 60, 2);
        //                            objoee.DownTimeBreakdown = Math.Round(DownTimeBreakdown, 2);
        //                            objoee.SummationOfSctvsPp = Math.Round(SummationOfSCTvsPP, 2);
        //                            objoee.ScrapQtyTime = Math.Round(ScrapQtyTime, 2);
        //                            objoee.ReWotime = Math.Round(ReWOTime, 2);
        //                            objoee.Loss1Name = lossCode1;
        //                            objoee.Loss1Value = lossCodeVal1;
        //                            objoee.Loss2Name = lossCode2;
        //                            objoee.Loss2Value = lossCodeVal2;
        //                            objoee.Loss3Name = lossCode3;
        //                            objoee.Loss3Value = lossCodeVal3;
        //                            objoee.Loss4Name = lossCode4;
        //                            objoee.Loss4Value = lossCodeVal4;
        //                            objoee.Loss5Name = lossCode5;
        //                            objoee.Loss5Value = lossCodeVal5;
        //                            objoee.CreatedOn = DateTime.Now;
        //                            objoee.CreatedBy = 1;
        //                            objoee.IsDeleted = 0;
        //                            dbLoss.Tbloeedashboardvariables.Add(objoee);
        //                            dbLoss.SaveChanges();

        //                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tbloeedashboardvariables (PlantID,ShopID,CellID,WCID,StartDate,EndDate,MinorLosses,Blue,Green,SettingTime,ROALossess,DownTimeBreakdown,SummationOfSCTvsPP,ScrapQtyTime,ReWOTime,Loss1Name,Loss1Value,Loss2Name,Loss2Value,Loss3Name,Loss3Value,Loss4Name,Loss4Value,Loss5Name,Loss5Value,CreatedOn,CreatedBy,IsDeleted)VALUES('" + PlantIDS + "','" + ShopIDS + "','" + CellIDS + "','" + MachineID + "','" + UsedDateForExcel.ToString("yyyy-MM-dd") + "','" + UsedDateForExcel.ToString("yyyy-MM-dd") + "','" + Math.Round(MinorLosses / 60, 2) + "','" + Math.Round(blue / 60, 2) + "','" + Math.Round(green / 60, 2) + "','" + Math.Round(SettingTime, 2) + "','" + Math.Round(ROALossess / 60, 2) + "','" + Math.Round(DownTimeBreakdown, 2) + "','" + Math.Round(SummationOfSCTvsPP, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReWOTime, 2) + "','" + lossCode1 + "','" + lossCodeVal1 + "','" + lossCode2 + "','" + lossCodeVal2 + "','" + lossCode3 + "','" + lossCodeVal3 + "','" + lossCode4 + "','" + lossCodeVal4 + "','" + lossCode5 + "','" + lossCodeVal5 + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + 1 + "','" + 0 + "');");

        //                        }
        //                        catch (Exception e)
        //                        {
        //                            result = false;
        //                        }
        //                        //finally
        //                        //{
        //                        //    mcInsertRows.close();
        //                        //}
        //                    }
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                result = false;
        //                //IntoFile("MacID: " + MachineID + e.ToString());
        //            }
        //        }
        //        result = true;
        //        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
        //    }
        //    #endregion
        //    return await Task.FromResult<bool>(result);
        //}

        //public async Task<double> GetMinorLosses(string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string dateString = currentdate.ToString("yyyy-MM-dd");

        //    double minorloss = 0;
        //    #region commented
        //    //int count = 0;
        //    //var Data = db.tbldailyprodstatus.Where(m => m.IsDeleted == 0 && m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).OrderBy(m => m.StartTime).ToList();
        //    //foreach (var row in Data)
        //    //{
        //    //    if (row.ColorCode == "yellow")
        //    //    {
        //    //        count++;
        //    //    }
        //    //    else
        //    //    {
        //    //        if (count > 0 && count < 2)
        //    //        {
        //    //            minorloss += count;
        //    //            count = 0;

        //    //        }
        //    //        count = 0;
        //    //    }
        //    //}

        //    #endregion
        //    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //    {
        //        var MinorLossSummation = dbLoss.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == dateString && m.ColorCode == Colour && m.DurationInSec < 120 && m.IsCompleted == 1).Sum(m => m.DurationInSec);
        //        minorloss = await Task.FromResult<double>(Convert.ToDouble(MinorLossSummation));
        //    }
        //    return minorloss;
        //}
        //public async Task<double> GetOPIDleBreakDown(string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");

        //    double count = 0;
        //    //MsqlConnection mc = new MsqlConnection();
        //    //mc.open();
        //    ////operating
        //    //mc.open();
        //    //String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //    //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //    //DataTable OP = new DataTable();
        //    //da1.Fill(OP);
        //    //mc.close();
        //    //if (OP.Rows.Count != 0)
        //    //{
        //    //    count[0] = Convert.ToInt32(OP.Rows[0][0]);
        //    //}

        //    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //    {
        //        var blah = dbLoss.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.ColorCode == Colour).Sum(m => m.DurationInSec);
        //        count = await Task.FromResult<double>(Convert.ToDouble(blah));
        //    }
        //    return count;
        //}

        //public async Task<double> GetSettingTime(string UsedDateForExcel, int MachineID)
        //{
        //    double settingTime = 0;
        //    int setupid = 0;
        //    string settingString = "Setup";
        //    var setupiddata = db.Tbllossescodes.Where(m => m.MessageType.Contains(settingString)).FirstOrDefault();
        //    if (setupiddata != null)
        //    {
        //        setupid = setupiddata.LossCodeId;
        //    }
        //    else
        //    {
        //        //Session["Error"] = "Unable to get Setup's ID";
        //        return -1;
        //    }
        //    // getting all setup's sublevels ids.
        //    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //    {
        //        var SettingIDs = dbLoss.Tbllossescodes.Where(m => m.LossCodesLevel1Id == setupid || m.LossCodesLevel2Id == setupid).Select(m => m.LossCodeId).ToList();


        //        //settingTime = (from row in db.tbllivelossofenties
        //        //where row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );


        //        var SettingData = dbLoss.Tbllivelossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
        //        foreach (var row in SettingData)
        //        {
        //            DateTime startTime = Convert.ToDateTime(row.StartDateTime);
        //            DateTime endTime = Convert.ToDateTime(row.EndDateTime);
        //            settingTime += endTime.Subtract(startTime).TotalMinutes;
        //        }
        //    }
        //    return await Task.FromResult<double>(settingTime);
        //}
        //public async Task<double> GetDownTimeLosses(string UsedDateForExcel, int MachineID, string contribute)
        //{
        //    double LossTime = 0;
        //    //string contribute = "ROA";
        //    // getting all ROA sublevels ids.Only those of IDLE.

        //    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //    {
        //        var SettingIDs = dbLoss.Tbllossescodes.Where(m => m.ContributeTo == contribute && (m.MessageType != "PM" || m.MessageType != "BREAKDOWN")).Select(m => m.LossCodeId).ToList();

        //        var SettingData = dbLoss.Tbllivelossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();

        //        var LossDuration = dbLoss.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.IsCompleted == 1 && m.DurationInSec > 120 && m.ColorCode == "YELLOW").Sum(m => m.DurationInSec);

        //        foreach (var row in SettingData)
        //        {
        //            DateTime startTime = Convert.ToDateTime(row.StartDateTime);
        //            DateTime endTime = Convert.ToDateTime(row.EndDateTime);
        //            LossTime += endTime.Subtract(startTime).TotalMinutes;
        //        }
        //        try
        //        {
        //            LossTime = (int)LossDuration;
        //        }
        //        catch { }
        //    }
        //    return await Task.FromResult<double>(LossTime);
        //}
        //public async Task<double> GetDownTimeBreakdown(string UsedDateForExcel, int MachineID)
        //{
        //    if (MachineID == 18)
        //    {
        //    }
        //    double LossTime = 0;
        //    using (i_facility_talContext dbLoss = new i_facility_talContext())
        //    {
        //        var BreakdownData = dbLoss.Tblbreakdown.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
        //        foreach (var row in BreakdownData)
        //        {
        //            if ((Convert.ToString(row.EndTime) == null) || row.EndTime == null)
        //            {
        //                //do nothing
        //            }
        //            else
        //            {
        //                DateTime startTime = Convert.ToDateTime(row.StartTime);
        //                DateTime endTime = Convert.ToDateTime(row.EndTime);
        //                LossTime += endTime.Subtract(startTime).TotalMinutes;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(LossTime);
        //}

        //public async Task<double> GetSummationOfSCTvsPP(string UsedDateForExcel, int MachineID)
        //{
        //    double SummationofTime = 0;
        //    //UsedDateForExcel = "2018-12-01";

        //    #region OLD 2017-02-10
        //    //var PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
        //    //if (PartsData.Count == 0)
        //    //{
        //    //    //return -1;
        //    //}
        //    //foreach (var row in PartsData)
        //    //{
        //    //    string partno = row.PartNo;
        //    //    string operationno = row.OperationNo;
        //    //    int totalpartproduced = Convert.ToInt32(row.DeliveredQty) + Convert.ToInt32(row.RejQty);
        //    //    Double stdCuttingTime = 0;
        //    //    var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
        //    //    if (stdcuttingTimeData != null)
        //    //    {
        //    //        string stdcuttingvalString = Convert.ToString(stdcuttingTimeData.StdCuttingTime);
        //    //        Double stdcuttingval = 0;
        //    //        if (double.TryParse(stdcuttingvalString, out stdcuttingval))
        //    //        {
        //    //            stdcuttingval = stdcuttingval;
        //    //        }

        //    //        string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //    //        if (Unit == "Hrs")
        //    //        {
        //    //            stdCuttingTime = stdcuttingval * 60;
        //    //        }
        //    //        else //Unit is Minutes
        //    //        {
        //    //            stdCuttingTime = stdcuttingval;
        //    //        }
        //    //    }
        //    //    SummationofTime += stdCuttingTime * totalpartproduced;
        //    //}
        //    ////To Extract MultiWorkOrder Cutting Time
        //    //PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWO == 1 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
        //    //if (PartsData.Count == 0)
        //    //{
        //    //    return SummationofTime;
        //    //}
        //    //foreach (var row in PartsData)
        //    //{
        //    //    int HMIID = row.HMIID;

        //    //    var DataInMultiwoSelection = db.tbl_multiwoselection.Where(m => m.HMIID == HMIID).ToList();
        //    //    foreach (var rowData in DataInMultiwoSelection)
        //    //    {
        //    //        string partno = rowData.PartNo;
        //    //        string operationno = rowData.OperationNo;
        //    //        int totalpartproduced = Convert.ToInt32(rowData.DeliveredQty) + Convert.ToInt32(rowData.ScrapQty);
        //    //        int stdCuttingTime = 0;
        //    //        var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
        //    //        if (stdcuttingTimeData != null)
        //    //        {
        //    //            int stdcuttingval = Convert.ToInt32(stdcuttingTimeData.StdCuttingTime);
        //    //            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //    //            if (Unit == "Hrs")
        //    //            {
        //    //                stdCuttingTime = stdcuttingval * 60;
        //    //            }
        //    //            else //Unit is Minutes
        //    //            {
        //    //                stdCuttingTime = stdcuttingval;
        //    //            }
        //    //        }
        //    //        SummationofTime += stdCuttingTime * totalpartproduced;
        //    //    }
        //    //}

        //    #endregion

        //    #region OLD 2017-02-10
        //    //List<string> OccuredWOs = new List<string>();
        //    ////To Extract Single WorkOrder Cutting Time
        //    //using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    //{
        //    //    var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.Hmiid).ToList();
        //    //    if (PartsDataAll.Count == 0)
        //    //    {
        //    //        //return SummationofTime;
        //    //    }
        //    //    foreach (var row in PartsDataAll)
        //    //    {
        //    //        string partNo = row.PartNo;
        //    //        string woNo = row.Work_Order_No;
        //    //        string opNo = row.OperationNo;

        //    //        string occuredwo = partNo + "," + woNo + "," + opNo;
        //    //        if (!OccuredWOs.Contains(occuredwo))
        //    //        {
        //    //            OccuredWOs.Add(occuredwo);
        //    //            var PartsData = dbhmi.Tblhmiscreen.
        //    //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0
        //    //                    && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)
        //    //                    && m.WorkOrderNo == woNo && m.PartNo == partNo && m.OperationNo == opNo).
        //    //                    OrderByDescending(m => m.Hmiid).ToList();

        //    //            int totalpartproduced = 0;
        //    //            int ProcessQty = 0, DeliveredQty = 0;
        //    //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

        //    //            #region new code

        //    //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
        //    //            int isHMIFirst = 2; //default NO History for that wo,pn,on

        //    //            var mulitwoData = dbhmi.TblMultiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWoid).Take(1).ToList();
        //    //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.IsWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

        //    //            //Note: we are in this loop => hmiscreen table data is Available

        //    //            if (mulitwoData.Count > 0)
        //    //            {
        //    //                isHMIFirst = 1;
        //    //            }
        //    //            else if (PartsData.Count > 0)
        //    //            {
        //    //                isHMIFirst = 0;
        //    //            }
        //    //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we both Dates now check for greatest amongst
        //    //            {
        //    //                int hmiIDFromMulitWO = row.HMIID;
        //    //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
        //    //                                                              where r.HMIID == hmiIDFromMulitWO
        //    //                                                              select r.Time
        //    //                                                              );
        //    //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

        //    //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
        //    //                {
        //    //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
        //    //                }
        //    //                else
        //    //                {
        //    //                    isHMIFirst = 0;
        //    //                }
        //    //            }
        //    //            if (isHMIFirst == 1)
        //    //            {
        //    //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
        //    //                int.TryParse(delivString, out DeliveredQty);
        //    //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
        //    //                int.TryParse(processString, out ProcessQty);

        //    //            }
        //    //            else if (isHMIFirst == 0)//Take Data from HMI
        //    //            {
        //    //                string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
        //    //                int.TryParse(delivString, out DeliveredQty);
        //    //                string processString = Convert.ToString(PartsData[0].ProcessQty);
        //    //                int.TryParse(processString, out ProcessQty);
        //    //            }

        //    //            #endregion

        //    //            //totalpartproduced = DeliveredQty + ProcessQty;
        //    //            totalpartproduced = DeliveredQty;

        //    //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

        //    //            double stdCuttingTime = 0;
        //    //            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //    //            if (stdcuttingTimeData != null)
        //    //            {
        //    //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //    //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //    //                if (Unit == "Hrs")
        //    //                {
        //    //                    stdCuttingTime = stdcuttingval * 60;
        //    //                }
        //    //                else //Unit is Minutes
        //    //                {
        //    //                    stdCuttingTime = stdcuttingval;
        //    //                }
        //    //            }
        //    //            #endregion

        //    //            SummationofTime += stdCuttingTime * totalpartproduced;
        //    //        }
        //    //    }
        //    //}
        //    ////To Extract Multi WorkOrder Cutting Time
        //    //using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    //{
        //    //    var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 1 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
        //    //    if (PartsDataAll.Count == 0)
        //    //    {
        //    //        //return SummationofTime;
        //    //    }
        //    //    foreach (var row in PartsDataAll)
        //    //    {
        //    //        string partNo = row.PartNo;
        //    //        string woNo = row.WorkOrderNo;
        //    //        string opNo = row.OperationNo;

        //    //        string occuredwo = partNo + "," + woNo + "," + opNo;
        //    //        if (!OccuredWOs.Contains(occuredwo))
        //    //        {
        //    //            OccuredWOs.Add(occuredwo);
        //    //            var PartsData = dbhmi.Tblhmiscreen.
        //    //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0
        //    //                    && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)
        //    //                    && m.WorkOrderNo == woNo && m.PartNo == partNo && m.OperationNo == opNo).
        //    //                    OrderByDescending(m => m.Hmiid).ToList();

        //    //            int totalpartproduced = 0;
        //    //            int ProcessQty = 0, DeliveredQty = 0;
        //    //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

        //    //            #region new code

        //    //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
        //    //            int isHMIFirst = 2; //default NO History for that wo,pn,on

        //    //            var mulitwoData = dbhmi.TblMultiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWoid).Take(1).ToList();
        //    //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.IsWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

        //    //            //Note: we are in this loop => hmiscreen table data is Available

        //    //            if (mulitwoData.Count > 0)
        //    //            {
        //    //                isHMIFirst = 1;
        //    //            }
        //    //            else if (PartsData.Count > 0)
        //    //            {
        //    //                isHMIFirst = 0;
        //    //            }
        //    //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we have both Dates now check for greatest amongst
        //    //            {
        //    //                int hmiIDFromMulitWO = row.Hmiid;
        //    //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
        //    //                                                              where r.HMIID == hmiIDFromMulitWO
        //    //                                                              select r.Time
        //    //                                                              );
        //    //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

        //    //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
        //    //                {
        //    //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
        //    //                }
        //    //                else
        //    //                {
        //    //                    isHMIFirst = 0;
        //    //                }
        //    //            }

        //    //            if (isHMIFirst == 1)
        //    //            {
        //    //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
        //    //                int.TryParse(delivString, out DeliveredQty);
        //    //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
        //    //                int.TryParse(processString, out ProcessQty);
        //    //            }
        //    //            else if (isHMIFirst == 0) //Take Data from HMI
        //    //            {
        //    //                string delivString = Convert.ToString(PartsData[0].DeliveredQty);
        //    //                int.TryParse(delivString, out DeliveredQty);
        //    //                string processString = Convert.ToString(PartsData[0].ProcessQty);
        //    //                int.TryParse(processString, out ProcessQty);
        //    //            }

        //    //            #endregion

        //    //            //totalpartproduced = DeliveredQty + ProcessQty;
        //    //            totalpartproduced = DeliveredQty;
        //    //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

        //    //            double stdCuttingTime = 0;
        //    //            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //    //            if (stdcuttingTimeData != null)
        //    //            {
        //    //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //    //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //    //                if (Unit == "Hrs")
        //    //                {
        //    //                    stdCuttingTime = stdcuttingval * 60;
        //    //                }
        //    //                else //Unit is Minutes
        //    //                {
        //    //                    stdCuttingTime = stdcuttingval;
        //    //                }
        //    //            }
        //    //            #endregion

        //    //            SummationofTime += stdCuttingTime * totalpartproduced;
        //    //        }
        //    //    }
        //    //}
        //    #endregion

        //    //new Code 2017-03-08
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
        //        if (PartsDataAll.Count == 0)
        //        {
        //            return SummationofTime;
        //        }
        //        foreach (var row in PartsDataAll)
        //        {
        //            if (row.IsMultiWo == 0)
        //            {
        //                string partNo = row.PartNo;
        //                string woNo = row.WorkOrderNo;
        //                string opNo = row.OperationNo;
        //                int DeliveredQty = 0;
        //                DeliveredQty = Convert.ToInt32(row.DeliveredQty);
        //                #region InnerLogic Common for both ways(HMI or tblmultiWOselection)
        //                double stdCuttingTime = 0;
        //                var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //                if (stdcuttingTimeData != null)
        //                {
        //                    double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //                    string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //                    if (Unit == "Hrs")
        //                    {
        //                        stdCuttingTime = stdcuttingval * 60;
        //                    }
        //                    else if (Unit == "Sec") //Unit is Minutes
        //                    {
        //                        stdCuttingTime = stdcuttingval / 60;
        //                    }
        //                    else
        //                    {
        //                        stdCuttingTime = stdcuttingval;
        //                    }
        //                    //no need of else , its already in minutes
        //                }
        //                #endregion
        //                //MessageBox.Show("CuttingTime " + stdCuttingTime + " DeliveredQty " + DeliveredQty);
        //                SummationofTime += stdCuttingTime * DeliveredQty;
        //                //MessageBox.Show("Single" + SummationofTime);
        //            }
        //            else
        //            {
        //                int hmiid = row.Hmiid;
        //                var multiWOData = dbhmi.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
        //                foreach (var rowMulti in multiWOData)
        //                {
        //                    string partNo = rowMulti.PartNo;
        //                    string opNo = rowMulti.OperationNo;
        //                    int DeliveredQty = 0;
        //                    DeliveredQty = Convert.ToInt32(rowMulti.DeliveredQty);
        //                    #region
        //                    double stdCuttingTime = 0;
        //                    var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
        //                    if (stdcuttingTimeData != null)
        //                    {
        //                        double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
        //                        string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
        //                        if (Unit == "Hrs")
        //                        {
        //                            stdCuttingTime = stdcuttingval * 60;
        //                        }
        //                        else if (Unit == "Sec") //Unit is Minutes
        //                        {
        //                            stdCuttingTime = stdcuttingval / 60;
        //                        }
        //                        else
        //                        {
        //                            stdCuttingTime = stdcuttingval;
        //                        }

        //                    }
        //                    #endregion
        //                    //MessageBox.Show("CuttingTime " + stdCuttingTime + " DeliveredQty " + DeliveredQty);
        //                    SummationofTime += stdCuttingTime * DeliveredQty;
        //                    //MessageBox.Show("Multi" + SummationofTime);
        //                }
        //            }
        //            //MessageBox.Show("" + SummationofTime);
        //        }
        //    }
        //    return await Task.FromResult<double>(SummationofTime);
        //}

        ////Output in Seconds
        //public async Task<double> GetGreen(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        //{
        //    double settingTime = 0;
        //    string stTime = StartTime.ToString("yyyy-MM-dd HH:mm:ss");
        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
        //          && m.ColorCode == "green" && m.StartTime <= StartTime && m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime) || (m.StartTime > StartTime && m.StartTime < EndTime)).ToList();
        //        foreach (var row in query1)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                settingTime += IndividualDur;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(settingTime);
        //}

        //public async Task<double> GetBlue(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        //{
        //    double settingTime = 0;

        //    using (i_facility_talContext db = new i_facility_talContext())
        //    {
        //        var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
        //          && m.ColorCode == "Blue" && m.StartTime <= StartTime && m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime) || (m.StartTime > StartTime && m.StartTime < EndTime)).ToList();

        //        foreach (var row in query1)
        //        {
        //            if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
        //            {
        //                DateTime LStartDate = Convert.ToDateTime(row.StartTime);
        //                DateTime LEndDate = Convert.ToDateTime(row.EndTime);
        //                double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

        //                //Get Duration Based on start & end Time.

        //                if (LStartDate < StartTime)
        //                {
        //                    double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
        //                    IndividualDur -= StartDurationExtra;
        //                }
        //                if (LEndDate > EndTime)
        //                {
        //                    double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
        //                    IndividualDur -= EndDurationExtra;
        //                }
        //                settingTime += IndividualDur;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(settingTime);
        //}

        //public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, int MachineID)
        //{
        //    double SQT = 0;
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsData = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 0).ToList();
        //        foreach (var row in PartsData)
        //        {
        //            string partno = row.PartNo;
        //            string operationno = row.OperationNo;
        //            int scrapQty = 0;
        //            int DeliveredQty = 0;
        //            string scrapQtyString = Convert.ToString(row.RejQty);
        //            string DeliveredQtyString = Convert.ToString(row.DeliveredQty);
        //            string x = scrapQtyString;
        //            int value;
        //            if (int.TryParse(x, out value))
        //            {
        //                scrapQty = value;
        //            }
        //            x = DeliveredQtyString;
        //            if (int.TryParse(x, out value))
        //            {
        //                DeliveredQty = value;
        //            }

        //            DateTime startTime = Convert.ToDateTime(row.Date);
        //            DateTime endTime = Convert.ToDateTime(row.Time);
        //            //Double WODuration = endTimeTemp.Subtract(startTime).TotalMinutes;
        //            Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

        //            if ((scrapQty + DeliveredQty) == 0)
        //            {
        //                SQT += 0;
        //            }
        //            else
        //            {
        //                SQT += ((WODuration / 60) / (scrapQty + DeliveredQty)) * scrapQty;
        //            }
        //        }
        //    }
        //    return await Task.FromResult<double>(SQT);
        //}

        ////GOD
        //public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, int MachineID)
        //{
        //    double SQT = 0;
        //    using (i_facility_talContext dbhmi = new i_facility_talContext())
        //    {
        //        var PartsData = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 1).ToList();
        //        foreach (var row in PartsData)
        //        {
        //            string partno = row.PartNo;
        //            string operationno = row.OperationNo;
        //            int scrapQty = Convert.ToInt32(row.RejQty);
        //            int DeliveredQty = Convert.ToInt32(row.DeliveredQty);
        //            DateTime startTime = Convert.ToDateTime(row.Date);
        //            DateTime endTime = Convert.ToDateTime(row.Time);
        //            Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

        //            //Double WODuration = endTime.Subtract(startTime).TotalMinutes;
        //            ////For Availability Loss
        //            //double Settingtime = GetSetupForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
        //            //double green = GetOT(UsedDateForExcel, startTime, endTime, MachineID);
        //            //double DownTime = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROA");
        //            //double BreakdownTime = GetBreakDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
        //            //double AL = DownTime + BreakdownTime + Settingtime;

        //            ////For Performance Loss
        //            //double downtimeROP = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROP");
        //            //double minorlossWO = GetMinorLossForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "yellow");
        //            //double PL = downtimeROP + minorlossWO;

        //            SQT += (WODuration / 60);
        //        }
        //    }
        //    return await Task.FromResult<double>(SQT);
        //}

        //#endregion End Oee Report

        //#endregion
        #endregion
    }
    #endregion

}
