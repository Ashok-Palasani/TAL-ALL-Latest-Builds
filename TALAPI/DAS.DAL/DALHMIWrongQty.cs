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
    public class DALHMIWrongQty : IHMIWrongQty
    {
        i_facility_talContext db = new i_facility_talContext();

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALHMIWrongQty));

        public static IConfiguration configuration;

        public DALHMIWrongQty(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }

        #region Prv Index

        //public CommonResponse IndexWOQty()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        List<WODetailsData> listWODetailsData = new List<WODetailsData>();
        //        string sentApprove = "";
        //        string acceptreject = "";
        //        string correctedDate = "";
        //        var getAllData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1).OrderBy(m=>m.CorrectedDate).ToList();
        //        if (getAllData.Count > 0)
        //        {
        //            foreach (var row in getAllData)
        //            {

        //                if (row.SendApprove == 1)
        //                {
        //                    sentApprove = "Sent For Approval";
        //                }
        //                if (row.AcceptReject == 1)
        //                {
        //                    acceptreject = "Yes";
        //                }
        //                else if (row.AcceptReject == 2)
        //                {
        //                    acceptreject = "No";
        //                }
        //                if(correctedDate !=row.CorrectedDate)
        //                {
        //                    correctedDate = row.CorrectedDate;
        //                    int machineId = Convert.ToInt32(row.MachineId);
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    WODetailsData objUnsignedWOData = new WODetailsData();
        //                    objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
        //                    objUnsignedWOData.MachineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.sendApprove = sentApprove;
        //                    objUnsignedWOData.acceptReject = acceptreject;
        //                    listWODetailsData.Add(objUnsignedWOData);
        //                }
        //                //WODetailsData objWODetailsData = new WODetailsData();
        //                //int machineID = row.MachineId;
        //                //string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
        //                //objWODetailsData.WoNo = row.WorkOrderNo;
        //                //objWODetailsData.OpNo = row.OperationNo;
        //                //objWODetailsData.partno = row.PartNo;
        //                //objWODetailsData.StartTime = Convert.ToDateTime(row.Date).ToString("yyyy-MM-dd HH:mm:ss");
        //                //objWODetailsData.EndTime = Convert.ToDateTime(row.Time).ToString("yyyy-MM-dd HH:mm:ss");
        //                //objWODetailsData.Project = row.Project;
        //                //objWODetailsData.prodfai = row.ProdFai;
        //                //objWODetailsData.WoQty = Convert.ToString(row.TargetQty);
        //                //objWODetailsData.ProcessQty = Convert.ToString(row.ProcessQty);
        //                //objWODetailsData.DeliveredQty = Convert.ToString(row.DeliveredQty);
        //                //objWODetailsData.WQtyhmiid = row.Wqtyhmiid;
        //                //objWODetailsData.MachineName = machineName;
        //                //objWODetailsData.sendApprove = sentApprove;
        //                //objWODetailsData.acceptReject = acceptreject;
        //                //listWODetailsData.Add(objWODetailsData);
        //            }
        //            obj.isTure = true;
        //            obj.response = listWODetailsData;
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
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        //public CommonResponse IndexWOQty()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        List<WODetailsData> listWODetailsData = new List<WODetailsData>();
        //        string firstApproval = "", secondApproval = "";
        //        string correctedDate = "";
        //        int machineid = 0;
        //        var getAllData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1).OrderByDescending(m => m.MachineId).ToList();
        //        if (getAllData.Count > 0)
        //        {
        //            foreach (var row in getAllData)
        //            {

        //                if (row.AcceptReject1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.AcceptReject1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.AcceptReject == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.AcceptReject == 2)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Rejected";
        //                }
        //                if (machineid != row.MachineId)
        //                {                            
        //                    correctedDate = row.CorrectedDate;
        //                    int machineId = Convert.ToInt32(row.MachineId);
        //                    machineid = machineId;
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    WODetailsData objUnsignedWOData = new WODetailsData();
        //                    objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
        //                    objUnsignedWOData.MachineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.firstApproval = firstApproval;
        //                    objUnsignedWOData.secondApproval = secondApproval;
        //                    listWODetailsData.Add(objUnsignedWOData);
        //                }

        //            }
        //            obj.isTure = true;
        //            obj.response = listWODetailsData;
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
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}
        #endregion

        //public CommonResponse IndexWOQty()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        List<WODetailsData> listWODetailsData = new List<WODetailsData>();

        //        string correctedDate = "";
        //        var getAllData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1).OrderBy(m => m.CorrectedDate).ToList();
        //        if (getAllData.Count > 0)
        //        {
        //            foreach (var row in getAllData)
        //            {
        //                string firstApproval = "", secondApproval = "";

        //                if (row.AcceptReject1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.AcceptReject1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.AcceptReject == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.AcceptReject == 2)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Rejected";
        //                }
        //                if (correctedDate != row.CorrectedDate)
        //                {
        //                    correctedDate = row.CorrectedDate;
        //                    int machineId = Convert.ToInt32(row.MachineId);
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    WODetailsData objUnsignedWOData = new WODetailsData();
        //                    objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
        //                    objUnsignedWOData.machineId = (int)row.MachineId;
        //                    objUnsignedWOData.MachineName = machineName.MachineInvNo;
        //                    objUnsignedWOData.plantName = plantName;
        //                    objUnsignedWOData.shopName = shopName;
        //                    objUnsignedWOData.cellName = cellName;
        //                    objUnsignedWOData.CorrectedDate = correctedDate;
        //                    objUnsignedWOData.firstApproval = firstApproval;
        //                    objUnsignedWOData.secondApproval = secondApproval;
        //                    listWODetailsData.Add(objUnsignedWOData);
        //                }
        //                else
        //                {
        //                    var listdata = listWODetailsData.Where(m => m.machineId == row.MachineId && m.CorrectedDate == row.CorrectedDate).FirstOrDefault();
        //                    if (listdata != null)
        //                    {

        //                    }
        //                    else
        //                    {
        //                        correctedDate = row.CorrectedDate;
        //                        int machineId = Convert.ToInt32(row.MachineId);
        //                        var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                        string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                        string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                        string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                        WODetailsData objUnsignedWOData = new WODetailsData();
        //                        objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
        //                        objUnsignedWOData.machineId = (int)row.MachineId;
        //                        objUnsignedWOData.MachineName = machineName.MachineInvNo;
        //                        objUnsignedWOData.plantName = plantName;
        //                        objUnsignedWOData.shopName = shopName;
        //                        objUnsignedWOData.cellName = cellName;
        //                        objUnsignedWOData.CorrectedDate = correctedDate;
        //                        objUnsignedWOData.firstApproval = firstApproval;
        //                        objUnsignedWOData.secondApproval = secondApproval;
        //                        listWODetailsData.Add(objUnsignedWOData);
        //                    }
        //                }

        //            }
        //            obj.isTure = true;
        //            obj.response = listWODetailsData;
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
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}


        //Index method for displaying



        //Get data for job finish and partial finsh data from tblhmiscreen

        public CommonResponse IndexWOQty()
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<WODetailsData> listWODetailsData = new List<WODetailsData>();

                string correctedDate = "";
                var getAllData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1).OrderByDescending(m => m.Wqtyhmiid).ToList();
                if (getAllData.Count > 0)
                {
                    foreach (var row in getAllData)
                    {
                        string firstApproval = "", secondApproval = "";

                        if (row.AcceptReject1 == 1)
                        {
                            secondApproval = "Aprroved";

                        }
                        else if (row.AcceptReject1 == 2)
                        {
                            secondApproval = "Rejected";

                        }
                        if (row.AcceptReject == 1)
                        {
                            firstApproval = "Aprroved";
                        }
                        else if (row.AcceptReject == 2)
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


                        //if (correctedDate != row.CorrectedDate)
                        //{
                        //    correctedDate = row.CorrectedDate;
                        //    int machineId = Convert.ToInt32(row.MachineId);
                        //    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
                        //    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
                        //    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
                        //    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
                        //    WODetailsData objUnsignedWOData = new WODetailsData();
                        //    objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
                        //    objUnsignedWOData.machineId = (int)row.MachineId;
                        //    objUnsignedWOData.MachineName = machineName.MachineInvNo;
                        //    objUnsignedWOData.plantName = plantName;
                        //    objUnsignedWOData.shopName = shopName;
                        //    objUnsignedWOData.cellName = cellName;
                        //    objUnsignedWOData.CorrectedDate = correctedDate;
                        //    objUnsignedWOData.firstApproval = firstApproval;
                        //    objUnsignedWOData.secondApproval = secondApproval;
                        //    listWODetailsData.Add(objUnsignedWOData);
                        //}
                        //else
                        //{
                        var listdata = listWODetailsData.Where(m => m.machineId == row.MachineId && m.CorrectedDate == row.CorrectedDate).FirstOrDefault();
                        if (listdata != null)
                        {

                        }
                        else
                        {
                            correctedDate = row.CorrectedDate;
                            int machineId = Convert.ToInt32(row.MachineId);
                            var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
                            string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
                            string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
                            string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
                            WODetailsData objUnsignedWOData = new WODetailsData();
                            objUnsignedWOData.WQtyhmiid = row.Wqtyhmiid;
                            objUnsignedWOData.machineId = (int)row.MachineId;
                            objUnsignedWOData.MachineName = machineName.MachineInvNo;
                            objUnsignedWOData.plantName = plantName;
                            objUnsignedWOData.shopName = shopName;
                            objUnsignedWOData.cellName = cellName;
                            objUnsignedWOData.CorrectedDate = correctedDate;
                            objUnsignedWOData.firstApproval = firstApproval;
                            objUnsignedWOData.secondApproval = secondApproval;
                            listWODetailsData.Add(objUnsignedWOData);
                        }
                        // }

                    }
                    obj.isTure = true;
                    obj.response = listWODetailsData;
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

        public CommonResponse GetWOJFPFDetails(Getdata data)
        {
            CommonResponse comobj = new CommonResponse();
            try
            {
                List<WODetails> listWODetails = new List<WODetails>();
                var wolist = db.Tblhmiscreen.Where(m => (m.Status == 2 || m.Status == 1) && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.WorkOrderNo == data.WoNo && m.OperationNo == data.OpNo).OrderBy(x => x.Hmiid).ToList();
                if (wolist.Count > 0)
                {
                    bool insertBool = InsertIntoBackupTable(wolist);
                    bool check = InsertIntoWQTYTable(wolist);
                    if (check)
                    {
                        var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == data.WoNo && x.OperationNo == data.OpNo).ToList();
                        if (wrongQtyData.Count > 0)
                        {
                            foreach (var rowwo in wrongQtyData)
                            {
                                int machineID = rowwo.MachineId;
                                bool jf = false, pf = false;
                                if (rowwo.IsWorkInProgress == 1)
                                {
                                    jf = true;
                                }
                                if (rowwo.IsWorkInProgress == 0)
                                {
                                    pf = true;
                                }
                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
                                WODetails objWODetails = new WODetails();
                                objWODetails.WoNo = rowwo.WorkOrderNo;
                                objWODetails.OpNo = rowwo.OperationNo;
                                objWODetails.partno = rowwo.PartNo;
                                objWODetails.StartTime = Convert.ToDateTime(rowwo.Date).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.EndTime = Convert.ToDateTime(rowwo.Time).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.Project = rowwo.Project;
                                objWODetails.prodfai = rowwo.ProdFai;
                                objWODetails.pf = pf;
                                objWODetails.jf = jf;
                                objWODetails.WoQty = Convert.ToString(rowwo.TargetQty);
                                objWODetails.ProcessQty = Convert.ToString(rowwo.ProcessQty);
                                objWODetails.DeliveredQty = Convert.ToString(rowwo.DeliveredQty);
                                objWODetails.WQtyhmiid = rowwo.Wqtyhmiid;
                                objWODetails.MachineName = machineName;
                                listWODetails.Add(objWODetails);
                            }
                            comobj.isTure = true;
                            comobj.response = listWODetails;
                        }
                        else
                        {
                            comobj.isTure = false;
                            comobj.response = ResourceResponse.NoItemsFound; ;
                        }
                    }
                    else
                    {
                        var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == data.WoNo && x.OperationNo == data.OpNo && x.SendApprove == 0).ToList();
                        if (wrongQtyData.Count > 0)
                        {
                            foreach (var rowwo in wrongQtyData)
                            {
                                int machineID = rowwo.MachineId;
                                bool jf = false, pf = false;
                                if (rowwo.IsWorkInProgress == 1)
                                {
                                    jf = true;
                                }
                                if (rowwo.IsWorkInProgress == 0)
                                {
                                    pf = true;
                                }
                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
                                WODetails objWODetails = new WODetails();
                                objWODetails.WoNo = rowwo.WorkOrderNo;
                                objWODetails.OpNo = rowwo.OperationNo;
                                objWODetails.partno = rowwo.PartNo;
                                objWODetails.StartTime = Convert.ToDateTime(rowwo.Date).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.EndTime = Convert.ToDateTime(rowwo.Time).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.Project = rowwo.Project;
                                objWODetails.prodfai = rowwo.ProdFai;
                                objWODetails.pf = pf;
                                objWODetails.jf = jf;
                                objWODetails.WoQty = Convert.ToString(rowwo.TargetQty);
                                objWODetails.ProcessQty = Convert.ToString(rowwo.ProcessQty);
                                objWODetails.DeliveredQty = Convert.ToString(rowwo.DeliveredQty);
                                objWODetails.WQtyhmiid = rowwo.Wqtyhmiid;
                                objWODetails.MachineName = machineName;
                                listWODetails.Add(objWODetails);
                            }
                            comobj.isTure = true;
                            comobj.response = listWODetails;
                        }
                        else
                        {
                            comobj.isTure = false;
                            comobj.response = ResourceResponse.NoItemsFound; ;
                        }
                    }
                }
                else
                {
                    comobj.isTure = false;
                    comobj.response = ResourceResponse.NoItemsFound;
                }

            }
            catch (Exception ex)
            {
                comobj.isTure = false;
                comobj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comobj;
        }

        public bool InsertIntoBackupTable(List<Tblhmiscreen> data)
        {
            bool ret = false;
            try
            {
                foreach (var rowhmi in data)
                {
                    var hmidet = db.TblBackUphmiscreen.Where(m => m.WorkOrderNo == rowhmi.WorkOrderNo && m.PartNo == rowhmi.PartNo && m.OperationNo == rowhmi.OperationNo && m.MachineId == rowhmi.MachineId).FirstOrDefault();
                    if (hmidet == null)
                    {
                        TblBackUphmiscreen tblhmi = new TblBackUphmiscreen();
                        tblhmi.BatchCount = rowhmi.BatchCount;
                        tblhmi.CorrectedDate = rowhmi.CorrectedDate;
                        tblhmi.Date = rowhmi.Date;
                        tblhmi.DdlwokrCentre = rowhmi.DdlwokrCentre;
                        tblhmi.DeliveredQty = rowhmi.DeliveredQty;
                        tblhmi.DoneWithRow = rowhmi.DoneWithRow;
                        tblhmi.Hmimonth = rowhmi.Hmimonth;
                        tblhmi.Hmiquarter = rowhmi.Hmiquarter;
                        tblhmi.HmiweekNumber = rowhmi.HmiweekNumber;
                        tblhmi.Hmiyear = rowhmi.Hmiyear;
                        tblhmi.IsHold = rowhmi.IsHold;
                        tblhmi.IsMultiWo = rowhmi.IsMultiWo;
                        tblhmi.IsSplitSapUpdated = rowhmi.IsSplitSapUpdated;
                        tblhmi.IsSync = rowhmi.IsSync;
                        tblhmi.IsUpdate = rowhmi.IsUpdate;
                        tblhmi.IsWorkInProgress = rowhmi.IsWorkInProgress;
                        tblhmi.IsWorkOrder = rowhmi.IsWorkOrder;
                        tblhmi.MachineId = rowhmi.MachineId;
                        tblhmi.OperationNo = Convert.ToString(rowhmi.OperatiorId);
                        tblhmi.OperatorDet = rowhmi.OperatorDet;
                        tblhmi.PartNo = rowhmi.PartNo;
                        tblhmi.PestartTime = rowhmi.PestartTime;
                        tblhmi.ProcessQty = rowhmi.ProcessQty;
                        tblhmi.ProdFai = rowhmi.ProdFai;
                        tblhmi.Project = rowhmi.Project;
                        tblhmi.RejQty = rowhmi.RejQty;
                        tblhmi.Shift = rowhmi.Shift;
                        tblhmi.SplitWo = rowhmi.SplitWo;
                        tblhmi.Status = rowhmi.Status;
                        tblhmi.TargetQty = rowhmi.TargetQty;
                        tblhmi.Time = rowhmi.Time;
                        tblhmi.WorkOrderNo = rowhmi.WorkOrderNo;
                        db.TblBackUphmiscreen.Add(tblhmi);
                        db.SaveChanges();
                        ret = true;
                    }
                }

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }


        #region Auto Suggest        

        public CommonResponse GetWorkOrderDetails(string workOrderNo)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<EntityHMIWrongQty> listEntityHMIWrongQty = new List<EntityHMIWrongQty>();
                DataTable dtWO = new DataTable();
                string workOrder = workOrderNo + "%";
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                SqlConnection conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand("AutoSuggestWorkOrder", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@workorder", workOrder);
                SqlDataAdapter sdaWO = new SqlDataAdapter(cmd);
                sdaWO.Fill(dtWO);
                if (dtWO.Rows.Count > 0)
                {
                    for (int i = 0; i < dtWO.Rows.Count; i++)
                    {
                        EntityHMIWrongQty objEntityHMIWrongQty = new EntityHMIWrongQty();
                        objEntityHMIWrongQty.workOrderNo = Convert.ToString(dtWO.Rows[i][0]);
                        listEntityHMIWrongQty.Add(objEntityHMIWrongQty);
                    }
                    obj.isTure = true;
                    obj.response = listEntityHMIWrongQty;
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


        //Get operation number
        public CommonResponse GetOperationDetails(AutoSuggestOpetaion data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<AutoSuggestOperationNo> listAutoSuggestOperationNo = new List<AutoSuggestOperationNo>();
                DataTable dtWO = new DataTable();
                string opNo = data.operationNo + "%";
                string workOrderNo = data.workOrderNo;
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                SqlConnection conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand("AutoSuggestOperationNo", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@operation", opNo);
                cmd.Parameters.AddWithValue("@workorder", workOrderNo);
                SqlDataAdapter sdaWO = new SqlDataAdapter(cmd);
                sdaWO.Fill(dtWO);
                if (dtWO.Rows.Count > 0)
                {
                    for (int i = 0; i < dtWO.Rows.Count; i++)
                    {
                        AutoSuggestOperationNo objAutoSuggestOperationNo = new AutoSuggestOperationNo();
                        objAutoSuggestOperationNo.operationNo = Convert.ToString(dtWO.Rows[i][0]);
                        listAutoSuggestOperationNo.Add(objAutoSuggestOperationNo);
                    }
                    obj.isTure = true;
                    obj.response = listAutoSuggestOperationNo;
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

        #endregion

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
                            var check1 = db.Tblwqtyhmiscreen.Where(m => intArry.Contains(m.Wqtyhmiid) && m.SendApprove == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check1.Count > 0)
                            {
                                var toMail1 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 5).Select(m => m.FirstApproverToList).FirstOrDefault();

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

                            var check2 = db.Tblwqtyhmiscreen.Where(m => intArry.Contains(m.Wqtyhmiid) && m.ApprovalLevel == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check2.Count > 0)
                            {
                                var toMail2 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 5).Select(m => m.SecondApproverToList).FirstOrDefault();
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
                obj.response = "Exception" + ex;
            }
            return obj;
        }

        // for updating the value and validating it
        public CommonResponse ValidateQtyUpdate(ValidateQTYUpdate data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int wQtyHMIId = data.WQtyhmiid;
                var wQtyData = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == wQtyHMIId).FirstOrDefault();
                if (wQtyData != null)
                {
                    if (wQtyData.Status == 2 && wQtyData.IsWorkInProgress == 1)
                    {
                        int deliveryQTY = Convert.ToInt32(data.qty);
                        int totalQty = Convert.ToInt32(wQtyData.TargetQty);
                        int processQty = Convert.ToInt32(wQtyData.ProcessQty);
                        int checkQty = deliveryQTY + processQty;
                        if (checkQty == totalQty)
                        {
                            wQtyData.DeliveredQty = deliveryQTY;
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                        else if (checkQty > totalQty)
                        {
                            obj.isTure = false;
                            obj.response = "The Qty " + deliveryQTY + " Entred is Greater Than Total Qty " + totalQty;
                        }
                        else if (checkQty < totalQty)
                        {
                            wQtyData.DeliveredQty = deliveryQTY;
                            wQtyData.Status = 1;
                            wQtyData.IsWorkInProgress = 0;
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                    }
                    else if (wQtyData.Status == 1 && wQtyData.IsWorkInProgress == 0)
                    {
                        int deliveryQTY = Convert.ToInt32(data.qty);
                        int processQty = Convert.ToInt32(wQtyData.ProcessQty);
                        int totalQty = Convert.ToInt32(wQtyData.TargetQty);
                        string workOrderNo = wQtyData.WorkOrderNo;
                        string operationNo = wQtyData.OperationNo;
                        string partNo = wQtyData.PartNo;
                        DateTime startTime = Convert.ToDateTime(wQtyData.Date);
                        var wQtyJFData = db.Tblwqtyhmiscreen.Where(x => x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.Status == 2 && x.IsWorkInProgress == 1).FirstOrDefault();
                        if (wQtyJFData == null)// if there no job finish row for particular wo,op and pn
                        {
                            var getAllRowWithIn = db.Tblwqtyhmiscreen.Where(x => x.Date > startTime && x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo).OrderBy(x => x.Date).ToList();
                            if (getAllRowWithIn.Count > 0)
                            {
                                int deliveryQtySum = deliveryQTY;
                                foreach (var row in getAllRowWithIn)
                                {
                                    deliveryQtySum += Convert.ToInt32(row.DeliveredQty);
                                }
                                if (deliveryQtySum == totalQty)
                                {
                                    int updationQty = deliveryQTY + processQty;
                                    foreach (var row in getAllRowWithIn)
                                    {
                                        row.ProcessQty = updationQty;
                                        db.SaveChanges();
                                        updationQty = Convert.ToInt32(row.DeliveredQty);
                                        if (updationQty == totalQty) // for last row update to JF
                                        {
                                            row.Status = 2;
                                            row.IsWorkInProgress = 1;
                                            db.SaveChanges();
                                        }
                                    }
                                    obj.isTure = true;
                                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                                }
                                else if (deliveryQtySum < totalQty)
                                {
                                    int updationQty = deliveryQTY + processQty;
                                    foreach (var row in getAllRowWithIn)
                                    {
                                        row.ProcessQty = updationQty;
                                        db.SaveChanges();
                                        updationQty = Convert.ToInt32(row.DeliveredQty);
                                    }
                                    obj.isTure = true;
                                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                                }
                                else if (deliveryQtySum > totalQty)
                                {
                                    obj.isTure = false;
                                    obj.response = "The Qty " + deliveryQtySum + " Entred is Greater Than Total Qty " + totalQty;
                                }
                            }
                            else
                            {
                                int checkQty = deliveryQTY + processQty;
                                if (checkQty == totalQty)
                                {
                                    wQtyData.DeliveredQty = deliveryQTY;
                                    wQtyData.Status = 2;
                                    wQtyData.IsWorkInProgress = 1;
                                    db.SaveChanges();
                                    obj.isTure = true;
                                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                                }
                                else if (checkQty < totalQty)
                                {
                                    wQtyData.DeliveredQty = deliveryQTY;
                                    db.SaveChanges();
                                    obj.isTure = true;
                                    obj.response = ResourceResponse.UpdatedSuccessMessage;

                                }
                                else if (checkQty > totalQty)
                                {
                                    obj.isTure = false;
                                    obj.response = "The Qty " + checkQty + " Entred is Greater Than Total Qty " + totalQty;
                                }
                            }
                        }
                        else  // if there job finish row for particular wo,op and pn
                        {
                            var getAllRowWithIn = db.Tblwqtyhmiscreen.Where(x => x.Date > startTime && x.PartNo == partNo && x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo).OrderBy(x => x.Date).ToList();
                            if (getAllRowWithIn.Count > 0)
                            {
                                int deliveryQtySum = deliveryQTY;
                                foreach (var row in getAllRowWithIn)
                                {
                                    deliveryQtySum += Convert.ToInt32(row.DeliveredQty);
                                }
                                if (deliveryQtySum == totalQty)
                                {
                                    int updationQty = deliveryQTY + processQty;
                                    foreach (var row in getAllRowWithIn)
                                    {
                                        row.ProcessQty = updationQty;
                                        db.SaveChanges();
                                        updationQty = Convert.ToInt32(row.DeliveredQty);
                                        if (updationQty == totalQty) // for last row update to JF
                                        {
                                            row.Status = 2;
                                            row.IsWorkInProgress = 1;
                                            db.SaveChanges();
                                            obj.isTure = true;
                                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                                        }
                                    }
                                }
                                else if (deliveryQtySum < totalQty)
                                {
                                    int updationQty = deliveryQTY + processQty;
                                    foreach (var row in getAllRowWithIn)
                                    {
                                        row.ProcessQty = updationQty;
                                        db.SaveChanges();
                                        updationQty = Convert.ToInt32(row.DeliveredQty);
                                    }
                                    obj.isTure = true;
                                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                                }
                                else if (deliveryQtySum > totalQty)
                                {
                                    obj.isTure = false;
                                    obj.response = "The Qty " + deliveryQtySum + " Entred is Greater Than Total Qty " + totalQty;
                                }
                            }
                            else
                            {
                                //obj.isTure = false;
                                //obj.response = ResourceResponse.NoItemsFound;    //// need to check
                            }
                        }
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

        //validate the quantity and with the selected workorder list
        public CommonResponse ValidateQtyData(GetIdsValues data)
        {
            CommonResponse obj = new CommonResponse();
            int index = 0, wQtyId = 0, value = 0, wQtyIdPrevoius = 0, valuePrevoius = 0, editedDelivertQty = 0, tblProcessQty = 0, previousDelivertQty = 0, previousProcessQty = 0, calProcessQty = 0, sumC = 0, sumP = 0;
            try
            {
                //getting the connection string from app string.json
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                string dbName = configuration.GetSection("MySettings").GetSection("Schema").Value;
                SqlConnection conn = new SqlConnection(connectionString);
                string truncateQuery = "truncate table " + dbName + ".[tblTempCalculate]";
                conn.Open();
                SqlCommand cmd = new SqlCommand(truncateQuery, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                string[] wQtyIds = data.wQtyHmiIds.Split(',');
                string[] values = data.values.Split(',');
                for (int i = 0; i < wQtyIds.Count(); i++)
                {
                    if (index == 0)
                    {
                        wQtyId = Convert.ToInt32(wQtyIds[i]);
                        value = Convert.ToInt32(values[i]);
                        var getValuesFromWQTY = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == wQtyId).FirstOrDefault();
                        if (getValuesFromWQTY != null)
                        {
                            editedDelivertQty = Convert.ToInt32(values[i]);
                            tblProcessQty = Convert.ToInt32(getValuesFromWQTY.ProcessQty);
                            TblTempCalculate addRow = new TblTempCalculate();
                            addRow.Whmiid = getValuesFromWQTY.Wqtyhmiid;
                            addRow.PdelQty = Convert.ToInt32(getValuesFromWQTY.DeliveredQty);
                            addRow.PprocQty = Convert.ToInt32(getValuesFromWQTY.ProcessQty);
                            addRow.CdelQty = editedDelivertQty;
                            addRow.CprocQty = tblProcessQty;
                            db.TblTempCalculate.Add(addRow);
                            db.SaveChanges();
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.NoItemsFound;
                            break;
                        }
                        index++;
                    }
                    else if (index == 1)
                    {
                        wQtyId = Convert.ToInt32(wQtyIds[i]);
                        value = Convert.ToInt32(values[i]);
                        wQtyIdPrevoius = Convert.ToInt32(wQtyIds[i - 1]);
                        valuePrevoius = Convert.ToInt32(values[i - 1]);
                        var getValuesFromWQTY = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == wQtyId).FirstOrDefault();
                        var getValuesFromWQTYPrevious = db.TblTempCalculate.Where(x => x.Whmiid == wQtyIdPrevoius).FirstOrDefault();
                        if (getValuesFromWQTY != null)
                        {
                            editedDelivertQty = Convert.ToInt32(values[i]);
                            previousDelivertQty = Convert.ToInt32(getValuesFromWQTYPrevious.CdelQty);
                            previousProcessQty = Convert.ToInt32(getValuesFromWQTYPrevious.CprocQty);
                            calProcessQty = previousDelivertQty + previousProcessQty;
                            TblTempCalculate addRow = new TblTempCalculate();
                            addRow.Whmiid = getValuesFromWQTY.Wqtyhmiid;
                            addRow.PdelQty = Convert.ToInt32(getValuesFromWQTY.DeliveredQty);
                            addRow.PprocQty = Convert.ToInt32(getValuesFromWQTY.ProcessQty);
                            addRow.CdelQty = editedDelivertQty;
                            addRow.CprocQty = calProcessQty;
                            db.TblTempCalculate.Add(addRow);
                            db.SaveChanges();
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.NoItemsFound;
                            break;
                        }
                    }
                }
                var validateQty = db.TblTempCalculate.Where(x => x.Whmiid != 0).OrderByDescending(x => x.Whmiid).FirstOrDefault();
                if (validateQty != null)
                {
                    sumC = Convert.ToInt32(validateQty.CdelQty) + Convert.ToInt32(validateQty.CprocQty);
                    sumP = Convert.ToInt32(validateQty.PdelQty) + Convert.ToInt32(validateQty.PprocQty);
                    if (sumC == sumP)
                    {
                        for (int i = 0; i < wQtyIds.Count(); i++)
                        {
                            wQtyId = Convert.ToInt32(wQtyIds[i]);
                            var getValuesFromWQTY = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == wQtyId).FirstOrDefault();
                            var getValuesFromWQTYValidate = db.TblTempCalculate.Where(x => x.Whmiid == wQtyId).FirstOrDefault();
                            if (getValuesFromWQTY != null)
                            {
                                getValuesFromWQTY.DeliveredQty = getValuesFromWQTYValidate.CdelQty;
                                getValuesFromWQTY.ProcessQty = getValuesFromWQTYValidate.CprocQty;
                                db.SaveChanges();
                            }
                        }
                        obj.isTure = true;
                        obj.response = ResourceResponse.SuccessMessage;
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = "The Sum OF Deliverty Qty :" + validateQty.CdelQty + " + Process Qty:" + validateQty.CprocQty + " Not Equal To Target Qty:" + sumP;
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

        //inserting rows in table
        public bool InsertIntoWQTYTable(List<Tblhmiscreen> hmiList)
        {
            bool result = false;
            try
            {
                if (hmiList.Count > 0)
                {
                    foreach (var row in hmiList)
                    {
                        var wrongQtyRow = db.Tblwqtyhmiscreen.Where(x => x.Bhmiid == row.Hmiid).FirstOrDefault();
                        if (wrongQtyRow == null)
                        {
                            InsertIntoWQTY(row);
                        }
                        else
                        {
                            //if (wrongQtyRow.SendApprove == 0) // delete previous data and insert new
                            //{
                            //    db.Tblwqtyhmiscreen.Remove(wrongQtyRow);
                            //    db.SaveChanges();
                            //    InsertIntoWQTY(row);
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //insert Medthod for wrong qty updation table
        public bool InsertIntoWQTY(Tblhmiscreen row)
        {
            bool result = false;
            int updateLevel = 1;
            try
            {
                var getMailIdsLevel = new List<TblTcfApprovedMaster>();

                var machData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => new { x.CellId, x.ShopId }).FirstOrDefault();
                int cellId = Convert.ToInt32(machData.CellId);
                int shopId = Convert.ToInt32(machData.ShopId);

                getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 5 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
                if (getMailIdsLevel.Count == 0)
                {
                    getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 5 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                }
                foreach (var rowMail in getMailIdsLevel)
                {
                    if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                    {
                        updateLevel = 2;
                    }
                }


                Tblwqtyhmiscreen addRow = new Tblwqtyhmiscreen();
                addRow.BatchCount = row.BatchCount;
                addRow.Bhmiid = row.Hmiid;
                addRow.CorrectedDate = row.CorrectedDate;
                addRow.Date = row.Date;
                addRow.DdlwokrCentre = row.DdlwokrCentre;
                addRow.DeliveredQty = row.DeliveredQty;
                addRow.DoneWithRow = row.DoneWithRow;
                addRow.Hmimonth = row.Hmimonth;
                addRow.Hmiquarter = row.Hmiquarter;
                addRow.HmiweekNumber = row.HmiweekNumber;
                addRow.Hmiyear = row.Hmiyear;
                addRow.IsHold = row.IsHold;
                addRow.IsMultiWo = row.IsMultiWo;
                addRow.IsSplitSapUpdated = row.IsSplitSapUpdated;
                addRow.IsSync = row.IsSync;
                addRow.IsUpdate = 0;
                addRow.IsWorkInProgress = row.IsWorkInProgress;
                addRow.IsWorkOrder = row.IsWorkOrder;
                addRow.Machine = row.Machine;
                addRow.MachineId = row.MachineId;
                addRow.OperationNo = row.OperationNo;
                addRow.OperatiorId = row.OperatiorId;
                addRow.OperatorDet = row.OperatorDet;
                addRow.PartNo = row.PartNo;
                addRow.PestartTime = row.PestartTime;
                addRow.ProcessQty = row.ProcessQty;
                addRow.ProdFai = row.ProdFai;
                addRow.Project = row.Project;
                addRow.RejQty = row.RejQty;
                addRow.Shift = row.Shift;
                addRow.SplitWo = row.SplitWo;
                addRow.Status = row.Status;
                addRow.TargetQty = row.TargetQty;
                addRow.Time = row.Time;
                addRow.WorkOrderNo = row.WorkOrderNo;
                addRow.UpdateLevel = updateLevel;
                db.Tblwqtyhmiscreen.Add(addRow);
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                result = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }


        // send mail on approval
        public CommonResponse SendApproval(Getdata data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string workOrderNo = data.WoNo;
                string operationNo = data.OpNo;
                int machineId = 0;
                //string correctedDate = data.Date;
                var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 0 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo).OrderBy(m => m.Wqtyhmiid).ToList();
                if (wrongQtyData.Count > 0)
                {
                 
                    string subjectName = "Wrong Qty " + workOrderNo + " " + operationNo;
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);
                    String[] seperator = { "{{WOStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{WOStart}}", "");
                    htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                    int sl = 1;
                    foreach (var row in wrongQtyData)
                    {
                        row.SendApprove = 1;
                        row.IsPending = 1;
                        db.SaveChanges();
                        machineId = row.MachineId;
                        String slno = Convert.ToString(sl);
                        int mchId = Convert.ToInt32(row.MachineId);
                        String operatorId = Convert.ToString(row.OperatiorId);
                        String targetQty = Convert.ToString(row.TargetQty);
                        String processQty = Convert.ToString(row.ProcessQty);
                        DateTime startDateTime = Convert.ToDateTime(row.Date);
                        DateTime endDateTime = Convert.ToDateTime(row.Time);
                        String deliveryQty = Convert.ToString(row.DeliveredQty);
                        String status = "";
                        if(row.Status == 2 && row.IsWorkInProgress == 1)
                        {
                            status = "Job Finish";
                        }
                        else if(row.Status == 1 && row.IsWorkInProgress == 0)
                        {
                            status = "Partial Finish";
                        }
                        else if(row.Remove == 1)
                        {
                            status = "Removed You can select for unassigned wo";
                        }
                        String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                        htmlStr = htmlStr.Replace("{{slno}}", slno);
                        htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                        htmlStr = htmlStr.Replace("{{Shift}}", row.Shift);
                        htmlStr = htmlStr.Replace("{{StartTime}}", startDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        htmlStr = htmlStr.Replace("{{EndTime}}", endDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                        htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                        htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                        htmlStr = htmlStr.Replace("{{OprationNo}}", row.OperationNo);
                        htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                        htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                        htmlStr = htmlStr.Replace("{{WorkOrderQty}}", targetQty);
                        htmlStr = htmlStr.Replace("{{ProcessedQty}}", processQty);
                        htmlStr = htmlStr.Replace("{{DeliveryQty}}", deliveryQty);
                        htmlStr = htmlStr.Replace("{{Status}}", status);

                        if (wrongQtyData.Count == 1)
                        {
                            htmlStr = htmlStr.Replace("{{WO}}", "");
                        }
                        else if (sl < wrongQtyData.Count)
                        {

                            htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                        }
                        else
                        {
                            htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                        }
                        sl++;
                       
                    }

                    htmlStr = htmlStr.Replace(woHtml, "");
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "For 1st Level Approval");

                    //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLWrongQty").Value;
                    string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLWrongQty").Value;


                    string rejectSrc = rejectUrl + "workOrderNo=" + workOrderNo + "&operationNo=" + operationNo + "&checked=0&MachineID=" + machineId + "";
                    //string acceptSrc = acceptUrl + "workOrderNo=" + workOrderNo + "&operationNo=" + operationNo + "";

                    string toName = "";
                    string toMailIds = "";
                    string ccMailIds = "";

                    var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 5).ToList();
                    foreach (var row in tcfApproveMail)
                    {
                        toMailIds += row.FirstApproverToList + ",";
                        ccMailIds += row.FirstApproverCcList + ",";
                    }


                    htmlStr = htmlStr.Replace("{{WO}}", "");
                    htmlStr = htmlStr.Replace("{{userName}}", "ALL");
                    //htmlStr = htmlStr.Replace("{{Sname}}", "Saurabh");
                    //htmlStr = htmlStr.Replace("{{Lurl}}", logo);
                    //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                    htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);



                    //string toMailID = "monika.ms@srkssolutions.com";
                    //string ccMailID = "vignesh.pai@srkssolutions.com,pavan.v@srkssolutions.com";
                    //string ccMailID = "vignesh.pai@srkssolutions.com,aswini.gp@srkssolutions.com";

                    /*string toMailID = toMailIds.Remove(toMailIds.Length - 1);*/// removing last comma
                    toMailIds = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                    ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);// removing last comma

                    bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 1, subjectName);

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

        // get data to show case on accept and rehect
        public CommonResponse GetAllData(GetMaildata data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                //bool validate= CheckApprovalLevel(data);
                //if (validate)
                //{ }
                //else
                //{
                string[] ids = data.id.Split(',');
                List<WODetails> listWODetails = new List<WODetails>();
                string workOrderNo = data.WoNo;
                string operationNo = data.OpNo;
                //string correctedDate = data.Date;
                if (data.id == "")
                {
                    //var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && (x.AcceptReject == 0 || x.AcceptReject == 1) && (x.AcceptReject1 == 0 || x.AcceptReject1 == 1)).ToList();
                    var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && x.IsPending == 1).ToList();
                    if (wrongQtyData.Count > 0)
                    {
                        foreach (var row in wrongQtyData)
                        {
                            int machineID = row.MachineId;
                            string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
                            WODetails objWODetails = new WODetails();
                            objWODetails.WoNo = row.WorkOrderNo;
                            objWODetails.OpNo = row.OperationNo;
                            objWODetails.partno = row.PartNo;
                            objWODetails.StartTime = Convert.ToDateTime(row.Date).ToString("yyyy-MM-dd HH:mm:ss");
                            objWODetails.EndTime = Convert.ToDateTime(row.Time).ToString("yyyy-MM-dd HH:mm:ss");
                            objWODetails.Project = row.Project;
                            objWODetails.prodfai = row.ProdFai;
                            objWODetails.WoQty = Convert.ToString(row.TargetQty);
                            objWODetails.ProcessQty = Convert.ToString(row.ProcessQty);
                            objWODetails.DeliveredQty = Convert.ToString(row.DeliveredQty);
                            objWODetails.WQtyhmiid = row.Wqtyhmiid;
                            objWODetails.MachineName = machineName;
                            listWODetails.Add(objWODetails);
                        }
                        obj.isTure = true;
                        obj.response = listWODetails;
                    }
                    else if (wrongQtyData.Count == 0)
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
                        var nocdet = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == id && x.IsPending == 0).FirstOrDefault();
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
                            var NoCodeDet = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == id && x.IsPending == 0 && (x.AcceptReject == 1 || x.AcceptReject1 == 1) && x.ApprovalLevel != 2).FirstOrDefault();
                            if (NoCodeDet != null)
                            {

                                //foreach (var row in NoCodeDet)
                                //{
                                int machineID = NoCodeDet.MachineId;
                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
                                WODetails objWODetails = new WODetails();
                                objWODetails.WoNo = NoCodeDet.WorkOrderNo;
                                objWODetails.OpNo = NoCodeDet.OperationNo;
                                objWODetails.partno = NoCodeDet.PartNo;
                                objWODetails.StartTime = Convert.ToDateTime(NoCodeDet.Date).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.EndTime = Convert.ToDateTime(NoCodeDet.Time).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.Project = NoCodeDet.Project;
                                objWODetails.prodfai = NoCodeDet.ProdFai;
                                objWODetails.WoQty = Convert.ToString(NoCodeDet.TargetQty);
                                objWODetails.ProcessQty = Convert.ToString(NoCodeDet.ProcessQty);
                                objWODetails.DeliveredQty = Convert.ToString(NoCodeDet.DeliveredQty);
                                objWODetails.WQtyhmiid = NoCodeDet.Wqtyhmiid;
                                objWODetails.MachineName = machineName;
                                listWODetails.Add(objWODetails);
                                //}
                                obj.isTure = true;
                                obj.response = listWODetails;
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
                            var NoCodeDet = db.Tblwqtyhmiscreen.Where(x => x.Wqtyhmiid == id && x.IsPending == 1).FirstOrDefault();
                            if (NoCodeDet != null)
                            {

                                //foreach (var row in NoCodeDet)
                                //{
                                int machineID = NoCodeDet.MachineId;
                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineID).Select(x => x.MachineInvNo).FirstOrDefault();
                                WODetails objWODetails = new WODetails();
                                objWODetails.WoNo = NoCodeDet.WorkOrderNo;
                                objWODetails.OpNo = NoCodeDet.OperationNo;
                                objWODetails.partno = NoCodeDet.PartNo;
                                objWODetails.StartTime = Convert.ToDateTime(NoCodeDet.Date).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.EndTime = Convert.ToDateTime(NoCodeDet.Time).ToString("yyyy-MM-dd HH:mm:ss");
                                objWODetails.Project = NoCodeDet.Project;
                                objWODetails.prodfai = NoCodeDet.ProdFai;
                                objWODetails.WoQty = Convert.ToString(NoCodeDet.TargetQty);
                                objWODetails.ProcessQty = Convert.ToString(NoCodeDet.ProcessQty);
                                objWODetails.DeliveredQty = Convert.ToString(NoCodeDet.DeliveredQty);
                                objWODetails.WQtyhmiid = NoCodeDet.Wqtyhmiid;
                                objWODetails.MachineName = machineName;
                                listWODetails.Add(objWODetails);
                                //}
                                obj.isTure = true;
                                obj.response = listWODetails;
                            }
                        }
                    }
                }
            }
            // }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        // Check the approval level and then update
        public bool CheckApprovalLevel(GetMaildata data)
        {
            bool result = false;
            try
            {
                int id = Convert.ToInt32(data.id);
                string workOrderNo = data.WoNo;
                string operationNo = data.OpNo;
                if (id == 0)
                {
                    var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.IsUpdate == 0 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && (x.AcceptReject == 2 || x.AcceptReject == 1)).ToList();
                    if (wrongQtyData.Count > 0)
                    {
                        result = true;
                    }
                }
                else if (id == -1)
                {
                    var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.IsUpdate == 0 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && (x.AcceptReject1 == 2 || x.AcceptReject1 == 1)).ToList();
                    if (wrongQtyData.Count > 0)
                    {
                        result = true;
                    }
                }
                else if (id == 1)
                {
                    var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.IsUpdate == 0 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && (x.AcceptReject == 2 || x.AcceptReject == 1)).ToList();
                    if (wrongQtyData.Count > 0)
                    {
                        result = true;
                    }
                }
                else if (id == 2)
                {
                    var wrongQtyData = db.Tblwqtyhmiscreen.Where(x => x.SendApprove == 1 && x.IsUpdate == 0 && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && (x.AcceptReject1 == 2 || x.AcceptReject1 == 1)).ToList();
                    if (wrongQtyData.Count > 0)
                    {
                        result = true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //Partial Finish
        public CommonResponsewithEror PartialFinish(pf data)
        {
            CommonResponsewithEror comobj = new CommonResponsewithEror();
            try
            {
                string[] wqtyids = data.wqtyids.Split(',').ToArray();
                for (int i = 0; i < wqtyids.Length; i++)
                {
                    int id = Convert.ToInt32(wqtyids[i]);
                    if (data.wqtyid == id)
                    {
                        var wrongqtydet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == data.wqtyid).FirstOrDefault();
                        if (wrongqtydet != null)
                        {

                            wrongqtydet.DeliveredQty = data.Delqty;
                            db.SaveChanges();
                            var prevdelqty = wrongqtydet.DeliveredQty;
                            var prevproqty = wrongqtydet.ProcessQty;
                            int Processedqty = (int)prevdelqty + prevproqty;
                            if (Processedqty < wrongqtydet.TargetQty)
                            {
                                wrongqtydet.Status = 1;
                                wrongqtydet.IsWorkInProgress = 0;
                                db.SaveChanges();

                                var ddldet = db.Tblddl.Where(m => m.WorkOrder == wrongqtydet.WorkOrderNo && m.OperationNo == wrongqtydet.OperationNo).FirstOrDefault();
                                if(ddldet!= null)
                                {
                                    if(ddldet.IsCompleted == 1)
                                    {
                                        ddldet.IsCompleted = 0;
                                        db.SaveChanges();
                                    }
                                }
                                comobj.isTure = true;
                                comobj.response = "Success";
                            }
                            else 
                            {
                                comobj.isTure = false;
                                comobj.errorMsg = "Delivered qty should be less than Target qty";
                                wrongqtydet.DeliveredQty = 0;
                                db.SaveChanges();
                            }
                        }
                    }
                   
                    else
                    {
                        if(id < data.wqtyid)
                        {
                            //break;
                        }
                        else
                        {
                            int id1 = Convert.ToInt32(wqtyids[i - 1]);
                            int nextid = Convert.ToInt32(wqtyids[i]);
                            var prewqtydet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == id1).FirstOrDefault();
                            if (prewqtydet != null)
                            {
                                var prevdelqty = prewqtydet.DeliveredQty;
                                var prevproqty = prewqtydet.ProcessQty;
                                int Processedqty = (int)prevdelqty + prevproqty;
                                var prewqtydet1 = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == nextid).FirstOrDefault();
                                if (prewqtydet1 != null)
                                {
                                    prewqtydet1.ProcessQty = Processedqty;
                                    db.SaveChanges();
                                    if (prewqtydet1.Status == 2 && prewqtydet1.IsWorkInProgress == 1)
                                    {
                                        int totalqty = (int)prewqtydet1.TargetQty;
                                        int newdelqty = totalqty - prewqtydet1.ProcessQty;
                                        prewqtydet1.DeliveredQty = newdelqty;
                                        db.SaveChanges();
                                    }
                                    else { }
                                    comobj.isTure = true;
                                    comobj.response = "Success";
                                }
                                if (prewqtydet1.TargetQty == Processedqty || prewqtydet1.TargetQty <= Processedqty)
                                {

                                }
                                //else if (Processedqty > prewqtydet1.TargetQty)
                                //{
                                //    int totalqty = (int)prewqtydet1.TargetQty;
                                //    int newdelqty = totalqty - prewqtydet1.ProcessQty;
                                //    prewqtydet1.DeliveredQty = newdelqty;
                                //    db.SaveChanges();
                                //    comobj.isTure = true;
                                //    comobj.response = "Success";
                                //}
                            }

                        }

                    }
                }

            }
            catch (Exception ex)
            {

            }
            return comobj;
        }

        //Job Finish
        public CommonResponsewithEror JobFinish(pf data)
        {
            CommonResponsewithEror comobj = new CommonResponsewithEror();
            try
            {
                string[] wqtyids = data.wqtyids.Split(',').ToArray();
                for (int i = 0; i < wqtyids.Length; i++)
                {
                    int id = Convert.ToInt32(wqtyids[i]);
                    if (data.wqtyid == id)
                    {
                        var wrongqtydet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == data.wqtyid).FirstOrDefault();
                        if (wrongqtydet != null)
                        {
                            wrongqtydet.DeliveredQty = data.Delqty;
                            db.SaveChanges();
                            var prevdelqty = wrongqtydet.DeliveredQty;
                            var prevproqty = wrongqtydet.ProcessQty;
                            int Processedqty = (int)prevdelqty + prevproqty;
                            if (Processedqty == wrongqtydet.TargetQty)
                            {
                                wrongqtydet.Status = 2;
                                wrongqtydet.IsWorkInProgress = 1;
                                db.SaveChanges();
                                comobj.isTure = true;
                                comobj.response = "Success";
                            }
                            else
                            {
                                comobj.isTure = false;
                                comobj.errorMsg = "Delivered qty Must be equal to Target qty";
                            }
                        }
                    }
                    else
                    {
                        if (id < data.wqtyid)
                        {

                        }
                        else
                        {
                            int nextid = Convert.ToInt32(wqtyids[i]);
                            var prewqtydet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == nextid).FirstOrDefault();
                            if (prewqtydet != null)
                            {
                                prewqtydet.Remove = 1;
                                db.SaveChanges();

                                var hmidet = db.Tblhmiscreen.Where(m => m.Hmiid == prewqtydet.Bhmiid).FirstOrDefault();
                                if (hmidet != null)
                                {
                                    db.Remove(hmidet);
                                    db.SaveChanges();
                                }
                                var livehmidet = db.Tbllivehmiscreen.Where(m => m.Hmiid == prewqtydet.Bhmiid).FirstOrDefault();
                                if (hmidet != null)
                                {
                                    db.Remove(livehmidet);
                                    db.SaveChanges();
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {

            }
            return comobj;
        }

        // accept the WorkOrder Updated QTY 
        public CommonResponse AcceptWoQtyData(Getdata data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string workOrderNo = data.WoNo;
                string operationNo = data.OpNo;
                // DateTime correctedDate = Convert.ToDateTime(data.Date);
                bool updateReport = false;
                int appLevel = 0;
                string subjectName = "";
                string[] ids = data.id.Split(',');
                int machineID = 0;
                foreach (var idrow in ids)
                {
                    int woqtyid = Convert.ToInt32(idrow);
                    var getWorkOrderQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.SendApprove == 1 && (x.AcceptReject == 0 || x.AcceptReject1 == 0) && x.Remove ==0 && x.Wqtyhmiid == woqtyid).FirstOrDefault();
                    if (getWorkOrderQtyData != null)
                    {
                        machineID = getWorkOrderQtyData.MachineId;
                        subjectName = "Wrong Qty " + workOrderNo + " " + operationNo;
                        //foreach (var row in getWorkOrderQtyData)
                        //{
                        if (getWorkOrderQtyData.AcceptReject == 0)
                        {
                            getWorkOrderQtyData.AcceptReject = 1;
                            getWorkOrderQtyData.IsPending = 0;
                            getWorkOrderQtyData.ApprovalLevel = 1;
                            db.SaveChanges();
                            if (getWorkOrderQtyData.UpdateLevel == 1)
                            {
                                updateReport = true;
                                appLevel = 1;
                            }
                            else
                            {
                                appLevel = 2;
                            }
                        }
                        else if (getWorkOrderQtyData.AcceptReject1 == 0)
                        {
                            getWorkOrderQtyData.AcceptReject1 = 1;
                            getWorkOrderQtyData.IsPending = 0;
                            getWorkOrderQtyData.ApprovalLevel = 2;
                            db.SaveChanges();
                            if (getWorkOrderQtyData.UpdateLevel == 2)
                            {
                                updateReport = true;
                            }
                            //appLevel = 2;
                        }
                        //}
                    }
                }
                if (data.unCheckId != "")
                {
                    string[] unCheckedids = data.unCheckId.Split(',');
                    foreach (var uncheckedIdRow in unCheckedids)
                    {
                        int id = Convert.ToInt32(uncheckedIdRow);
                        var getNoCodeDet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == id).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            getNoCodeDet.IsPending = 1;
                            db.SaveChanges();
                        }
                    }
                }
                string toMailIds = "";
                string ccMailIds = "";
                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 5).ToList();
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

                toMailIds = toMailIds.Remove(toMailIds.Length - 1);
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);

                if (updateReport)
                {
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);
                    String[] seperator = { "{{WOStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{WOStart}}", "");
                    htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                    int sl = 1;
                    var getWorkOrderQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.SendApprove == 1 && x.AcceptReject == 1 && x.AcceptReject1 == 1 && x.IsUpdate == 0 && x.Remove == 0).OrderBy(m => m.Wqtyhmiid).ToList();

                    if (getWorkOrderQtyData.Count > 0)
                    {
                        foreach (var row in getWorkOrderQtyData)
                        {

                            row.SendApprove = 1;
                            db.SaveChanges();

                            String slno = Convert.ToString(sl);
                            int mchId = Convert.ToInt32(row.MachineId);
                            String operatorId = Convert.ToString(row.OperatiorId);
                            String targetQty = Convert.ToString(row.TargetQty);
                            String processQty = Convert.ToString(row.ProcessQty);
                            DateTime startDateTime = Convert.ToDateTime(row.Date);
                            DateTime endDateTime = Convert.ToDateTime(row.Time);
                            String deliveryQty = Convert.ToString(row.DeliveredQty);
                            String status = "";
                            if (row.Status == 2 && row.IsWorkInProgress == 1)
                            {
                                status = "Job Finish";
                            }
                            else if (row.Status == 1 && row.IsWorkInProgress == 0)
                            {
                                status = "Partial Finish";
                            }
                            else if (row.Remove == 1)
                            {
                                status = "Removed You can select for unassigned wo";
                            }
                            String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{Shift}}", row.Shift);
                            htmlStr = htmlStr.Replace("{{StartTime}}", startDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{EndTime}}", endDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.OperationNo);
                            htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                            htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                            htmlStr = htmlStr.Replace("{{WorkOrderQty}}", targetQty);
                            htmlStr = htmlStr.Replace("{{ProcessedQty}}", processQty);
                            htmlStr = htmlStr.Replace("{{DeliveryQty}}", deliveryQty);
                            htmlStr = htmlStr.Replace("{{Status}}", status);

                            if (getWorkOrderQtyData.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < getWorkOrderQtyData.Count)
                            {

                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            sl++;
                        }
                        htmlStr = htmlStr.Replace(woHtml, "");
                        htmlStr = htmlStr.Replace("{{secondLevel}}", "Thease Work Order are Accepted");

                        bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 2, subjectName);

                        if (ret)
                        {
                            obj.isTure = true;
                            obj.response = "Sent Mail";
                            UpdateToReportTables(workOrderNo, operationNo);
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.FailureMessage;
                        }
                    }

                    // string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                    //"</head><body><p>Dear ALL" + ",</p></br><p><center> The Wrong QtY updation Has Been Accepted</center></p></br><p>Thank you" +
                    //"</p></body></html>";

                    // bool ret = SendMail(message, toMailIds, ccMailIds, 0,subjectName);

                    // if (ret)
                    // {
                    //     UpdateToReportTables(workOrderNo, operationNo);
                    //     obj.isTure = true;
                    //     obj.response = "Data Updated in Reports";
                    // }
                    // else
                    // {
                    //     obj.isTure = false;
                    //     obj.response = ResourceResponse.FailureMessage;
                    // }
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

                    int sl = 1;
                    var getWorkOrderQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.SendApprove == 1 && x.AcceptReject == 1 && x.AcceptReject1 == 0 && x.IsUpdate == 0 && x.Remove == 0).OrderBy(m => m.Wqtyhmiid).ToList();

                    if (getWorkOrderQtyData.Count > 0)
                    {
                        foreach (var row in getWorkOrderQtyData)
                        {

                            //row.SendApprove = 1;
                            //db.SaveChanges();

                            String slno = Convert.ToString(sl);
                            int mchId = Convert.ToInt32(row.MachineId);
                            String operatorId = Convert.ToString(row.OperatiorId);
                            String targetQty = Convert.ToString(row.TargetQty);
                            String processQty = Convert.ToString(row.ProcessQty);
                            DateTime startDateTime = Convert.ToDateTime(row.Date);
                            DateTime endDateTime = Convert.ToDateTime(row.Time);
                            String deliveryQty = Convert.ToString(row.DeliveredQty);
                            String status = "";
                            if (row.Status == 2 && row.IsWorkInProgress == 1)
                            {
                                status = "Job Finish";
                            }
                            else if (row.Status == 1 && row.IsWorkInProgress == 0)
                            {
                                status = "Partial Finish";
                            }
                            else if (row.Remove == 1)
                            {
                                status = "Removed You can select for unassigned wo";
                            }
                            String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{Shift}}", row.Shift);
                            htmlStr = htmlStr.Replace("{{StartTime}}", startDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{EndTime}}", endDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.OperationNo);
                            htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                            htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                            htmlStr = htmlStr.Replace("{{WorkOrderQty}}", targetQty);
                            htmlStr = htmlStr.Replace("{{ProcessedQty}}", processQty);
                            htmlStr = htmlStr.Replace("{{DeliveryQty}}", deliveryQty);
                            htmlStr = htmlStr.Replace("{{Status}}", status);

                            if (getWorkOrderQtyData.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < getWorkOrderQtyData.Count)
                            {

                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                            }
                            sl++;
                        }

                        htmlStr = htmlStr.Replace(woHtml, "");
                        htmlStr = htmlStr.Replace("{{secondLevel}}", "For 2nd Level Approval");

                        //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLWrongQty").Value;
                        string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLWrongQty").Value;

                        //acceptUrl = acceptUrl.Replace("id=1", "id=2");  // Work disabling the button on the mail
                        //rejectUrl = rejectUrl.Replace("id=0", "id=-1");


                        string rejectSrc = rejectUrl + "workOrderNo=" + workOrderNo + "&operationNo=" + operationNo + "&checked=" + data.id + "&MachineID="+machineID+"";
                        // string acceptSrc = acceptUrl + "workOrderNo=" + workOrderNo + "&operationNo=" + operationNo + "";

                        htmlStr = htmlStr.Replace("{{WO}}", "");
                        htmlStr = htmlStr.Replace("{{userName}}", "ALL");

                        //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                        htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);



                        bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 1, subjectName);

                        if (ret)
                        {
                            obj.isTure = true;
                            obj.response = "Sent Mail for Second Level Approval";
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

                // }
                //else
                //{
                //    obj.isTure = false;
                //    obj.response = ResourceResponse.NoItemsFound;
                //}
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        // update the record to live tables
        public bool UpdateQtyLiveHist(Tblwqtyhmiscreen data)
        {
            bool result = false;
            try
            {
                int hmiId = data.Bhmiid;
                string workOrderNo = data.WorkOrderNo;
                string operationNo = data.OperationNo;
                string correctedDate = data.CorrectedDate;
                int isWorkInProgress = data.IsWorkInProgress;
                int status = Convert.ToInt32(data.Status);
                int deliveryQty = Convert.ToInt32(data.DeliveredQty);
                int processQty = Convert.ToInt32(data.ProcessQty);
                var checkHmiDataUpdate = db.Tblhmiscreen.Where(x => x.CorrectedDate == correctedDate && x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.Hmiid == hmiId).FirstOrDefault();
                if (checkHmiDataUpdate != null)
                {
                    checkHmiDataUpdate.DeliveredQty = deliveryQty;
                    //checkHmiDataUpdate.Status = status;
                    //checkHmiDataUpdate.IsWorkInProgress = isWorkInProgress;
                    checkHmiDataUpdate.ProcessQty = processQty;
                    db.SaveChanges();
                    var liveHMIdata = db.Tbllivehmiscreen.Where(x => x.CorrectedDate == correctedDate && x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.Hmiid == hmiId).FirstOrDefault();
                    if (liveHMIdata != null)
                    {
                        liveHMIdata.DeliveredQty = deliveryQty;
                        //liveHMIdata.Status = status;
                        //liveHMIdata.IsWorkInProgress = isWorkInProgress;
                        checkHmiDataUpdate.ProcessQty = processQty;
                        db.SaveChanges();
                    }
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                result = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //Get the reject Reasons
        public CommonResponse GetRejectReasons()
        {
            CommonResponse obj = new CommonResponse();
            try
            {
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
            }
            catch (Exception ex)
            {

            }
            return obj;
        }


        // accept the WorkOrder Updated QTY 
        public CommonResponse RejectWoQtyData(GetdataWithRejectReason data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int checkMail = 0;
                string workOrderNo = data.WoNo;
                string operationNo = data.OpNo;
                string subjectName = "";
                //DateTime correctedDate = Convert.ToDateTime(data.Date);
                string[] ids = data.id.Split(',');
                foreach (var idrow in ids)
                {
                    int woqtyid = Convert.ToInt32(idrow);
                    var getWorkOrderQtyData = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.SendApprove == 1 && (x.AcceptReject == 0 || x.AcceptReject1 == 0) && x.Wqtyhmiid == woqtyid && x.Remove == 0).FirstOrDefault();
                    if (getWorkOrderQtyData != null)
                    {
                        subjectName = "Wrong Qty " + workOrderNo + " " + operationNo;
                        //foreach (var row in getWorkOrderQtyData)
                        //{                     
                        if (getWorkOrderQtyData.AcceptReject == 0)
                        {
                            getWorkOrderQtyData.AcceptReject = 2;
                            getWorkOrderQtyData.IsPending = 0;
                            getWorkOrderQtyData.RejectReasonId = data.reasonId;
                            getWorkOrderQtyData.ApprovalLevel = 1;
                            db.SaveChanges();
                            checkMail = 1;
                        }
                        else if (getWorkOrderQtyData.AcceptReject1 == 0)
                        {
                            getWorkOrderQtyData.AcceptReject1 = 2;
                            getWorkOrderQtyData.IsPending = 0;
                            getWorkOrderQtyData.RejectReason1 = data.reasonId;
                            getWorkOrderQtyData.ApprovalLevel = 2;
                            db.SaveChanges();
                            checkMail = 2;
                        }

                    }
                }
                if (data.unCheckId != "")
                {
                    string[] unCheckedids = data.unCheckId.Split(',');
                    foreach (var uncheckedIdRow in unCheckedids)
                    {
                        int id = Convert.ToInt32(uncheckedIdRow);
                        var getNoCodeDet = db.Tblwqtyhmiscreen.Where(m => m.Wqtyhmiid == id).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            getNoCodeDet.IsPending = 1;
                            db.SaveChanges();
                        }
                    }
                }
                string toMailIds = "";
                string ccMailIds = "";
                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 5).ToList();
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
                toMailIds = toMailIds.Remove(toMailIds.Length - 1);
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);

                string rejectName = db.Tblrejectreason.Where(x => x.IsDeleted == 0 && x.Rid == data.reasonId).Select(x => x.RejectNameDesc).FirstOrDefault();

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\WOAceptRejectTemplate1.html");
                string htmlStr = File.ReadAllText(reader);
                String[] seperator = { "{{WOStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{WOStart}}", "");
                htmlStr = htmlStr.Replace("{{WOEnd}}", "");

                int sl = 1;
                var getWorkOrderQtyData1 = db.Tblwqtyhmiscreen.Where(x => x.WorkOrderNo == workOrderNo && x.OperationNo == operationNo && x.SendApprove == 1 && x.AcceptReject == 2 && x.AcceptReject1 == 2 && x.IsUpdate == 0 && x.Remove == 0).OrderBy(m => m.Wqtyhmiid).ToList();

                if (getWorkOrderQtyData1.Count > 0)
                {
                    foreach (var row in getWorkOrderQtyData1)
                    {

                        row.SendApprove = 1;
                        db.SaveChanges();

                        String slno = Convert.ToString(sl);
                        int mchId = Convert.ToInt32(row.MachineId);
                        String operatorId = Convert.ToString(row.OperatiorId);
                        String targetQty = Convert.ToString(row.TargetQty);
                        String processQty = Convert.ToString(row.ProcessQty);
                        DateTime startDateTime = Convert.ToDateTime(row.Date);
                        DateTime endDateTime = Convert.ToDateTime(row.Time);
                        String deliveryQty = Convert.ToString(row.DeliveredQty);
                        String status = "";
                        if (row.Status == 2 && row.IsWorkInProgress == 1)
                        {
                            status = "Job Finish";
                        }
                        else if (row.Status == 1 && row.IsWorkInProgress == 0)
                        {
                            status = "Partial Finish";
                        }
                        else if (row.Remove == 1)
                        {
                            status = "Removed You can select for unassigned wo";
                        }
                        String machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == mchId).Select(x => x.MachineInvNo).FirstOrDefault();
                        htmlStr = htmlStr.Replace("{{slno}}", slno);
                        htmlStr = htmlStr.Replace("{{MachineName}}", machineName);
                        htmlStr = htmlStr.Replace("{{Shift}}", row.Shift);
                        htmlStr = htmlStr.Replace("{{StartTime}}", startDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        htmlStr = htmlStr.Replace("{{EndTime}}", endDateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        htmlStr = htmlStr.Replace("{{OperatorId}}", operatorId);
                        htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.WorkOrderNo);
                        htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNo);
                        htmlStr = htmlStr.Replace("{{OprationNo}}", row.OperationNo);
                        htmlStr = htmlStr.Replace("{{Project}}", row.Project);
                        htmlStr = htmlStr.Replace("{{ProdFAI}}", row.ProdFai);
                        htmlStr = htmlStr.Replace("{{WorkOrderQty}}", targetQty);
                        htmlStr = htmlStr.Replace("{{ProcessedQty}}", processQty);
                        htmlStr = htmlStr.Replace("{{DeliveryQty}}", deliveryQty);
                        htmlStr = htmlStr.Replace("{{Status}}", status);

                        if (getWorkOrderQtyData1.Count == 1)
                        {
                            htmlStr = htmlStr.Replace("{{WO}}", "");
                        }
                        else if (sl < getWorkOrderQtyData1.Count)
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
                htmlStr = htmlStr.Replace(woHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "Thease Work Order Has Been Rejected for this " + rejectName + ".");


                // string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                //"</head><body><p>Dear All,</p></br><p><center> The Wrong Qty Updation Has Been Rejected for this " + rejectName + ".</center></p></br><p>Thank you" +
                //"</p></br></body></html>";



                bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 2, subjectName);

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
                //}
                //else
                //{
                //    obj.isTure = false;
                //    obj.response = ResourceResponse.NoItemsFound;
                //}
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        public bool SendMail(string message, string toList, string ccList, int image, string subject)
        {
            bool ret = false;
            try
            {
                if (message != "" && toList != "" && ccList != null)
                {
                    string toMailID = toList;
                    string ccMailID = ccList;
                    MailMessage mail = new MailMessage();
                    mail.To.Add(toMailID);
                    if (ccMailID != "")
                    {
                        mail.CC.Add(ccMailID);
                    }

                    var smtpConn = db.Smtpdetails.Where(x => x.IsDeleted == true && x.TcfModuleId == 5).FirstOrDefault();
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


                    //string fromMail = configuration.GetSection("SMTPConn").GetSection("FromMailID").Value;

                    mail.From = new MailAddress(fromMail);
                    mail.Subject = subject;
                    mail.Body = "" + message;
                    mail.IsBodyHtml = true;
                    mail.BodyEncoding = System.Text.Encoding.UTF8;
                    mail.SubjectEncoding = System.Text.Encoding.UTF8;

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
                        // htmlView.LinkedResources.Add(img2);
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
                    log.Error("SendMailSuccessfully");

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

        //getting all the record for updation into table
        public CommonResponse UpdateToReportTables(string woNo, string OpNo)
        {
            bool check = false;
            string correctedDate = "";
            CommonResponse obj = new CommonResponse();
            try
            {
                var unSignedWO = db.Tblwqtyhmiscreen.Where(x => x.IsUpdate == 0 && (x.AcceptReject == 1 || x.AcceptReject1 == 1) && x.OperationNo == OpNo && x.WorkOrderNo == woNo).OrderBy(m => m.Wqtyhmiid).ToList();
                foreach (var row in unSignedWO)
                {
                    correctedDate = row.CorrectedDate;
                    check = UpdateQtyLiveHist(row);
                }

                if (check)
                {
                    check = TakeBackupReportData(correctedDate);
                    if (check)
                    {
                        DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
                        DateTime correcteDateTime = Convert.ToDateTime(correctedDate);
                        List<int> machinelist = new List<int>();
                        machinelist = db.Tblwqtyhmiscreen.Where(m => m.WorkOrderNo == woNo && m.OperationNo == OpNo).Select(m => m.MachineId).ToList();
                        Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(correcteDateTime, correcteDateTime,machinelist);  // for WO report updation
                        Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(correcteDateTime, correcteDateTime,machinelist);// for OEE report updation
                        if (reportOEEUpdate.Result == true && reportOEEUpdate.Result == true)
                        {
                            foreach (var row in unSignedWO)
                            {
                                row.IsUpdate = 1;
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

        #region Insert method

        //insert method for the report table and related tables
        //public bool InsertToLiveHMI(Tblwqtyhmiscreen data)
        //{
        //    bool ret = false;
        //    try
        //    {
        //        Tbllivehmiscreen addRow = new Tbllivehmiscreen();
        //        addRow.CorrectedDate = data.CorrectedDate;
        //        addRow.Date = Convert.ToDateTime(data.Date);
        //        addRow.DeliveredQty = Convert.ToInt32(data.DeliveredQty);
        //        addRow.IsWorkInProgress = Convert.ToInt32(data.IsWorkInProgress);
        //        addRow.MachineId = Convert.ToInt32(data.MachineId);
        //        addRow.OperationNo = data.OperationNo;
        //        addRow.OperatorDet = Convert.ToString(data.OperatiorId);
        //        addRow.PartNo = data.PartNo;
        //        addRow.ProcessQty = Convert.ToInt32(data.ProcessQty);
        //        addRow.ProdFai = data.ProdFai;
        //        addRow.Project = data.Project;
        //        addRow.Shift = data.Shift;
        //        addRow.Status = data.Status;
        //        addRow.TargetQty = Convert.ToInt32(data.TargetQty);
        //        addRow.Time = Convert.ToDateTime(data.Time);
        //        addRow.WorkOrderNo = data.WorkOrderNo;
        //        addRow.DdlwokrCentre = data.DdlwokrCentre;
        //        addRow.IsWorkOrder = Convert.ToInt32(data.IsWorkOrder);
        //        db.Tbllivehmiscreen.Add(addRow);
        //        db.SaveChanges();

        //        int hmiid = addRow.Hmiid;

        //        Tblhmiscreen addRowHmi = new Tblhmiscreen();
        //        addRowHmi.Hmiid = hmiid;
        //        addRowHmi.CorrectedDate = data.CorrectedDate;
        //        addRowHmi.Date = Convert.ToDateTime(data.Date);
        //        addRowHmi.DeliveredQty = Convert.ToInt32(data.DeliveredQty);
        //        addRowHmi.IsWorkInProgress = 1;
        //        addRowHmi.MachineId = Convert.ToInt32(data.MachineId);
        //        addRowHmi.OperationNo = data.OperationNo;
        //        addRowHmi.OperatorDet = Convert.ToString(data.OperatiorId);
        //        addRowHmi.PartNo = data.PartNo;
        //        addRowHmi.ProcessQty = Convert.ToInt32(data.ProcessQty);
        //        addRowHmi.ProdFai = data.ProdFai;
        //        addRowHmi.Project = data.Project;
        //        addRowHmi.Shift = data.Shift;
        //        addRowHmi.Status = 2;
        //        addRowHmi.TargetQty = Convert.ToInt32(data.TargetQty);
        //        addRowHmi.Time = Convert.ToDateTime(data.Time);
        //        addRowHmi.WorkOrderNo = data.WorkOrderNo;
        //        addRow.DdlwokrCentre = data.DdlwokrCentre;
        //        db.Tblhmiscreen.Add(addRowHmi);
        //        db.SaveChanges();

        //        if (data.IsWorkInProgress == 1 && data.Status == 2)
        //        {
        //            var updateRow = db.Tblddl.Where(x => x.IsDeleted == 0 && x.WorkOrder == data.WorkOrderNo && x.MaterialDesc == data.PartNo && x.OperationNo == data.OperationNo).FirstOrDefault();
        //            if (updateRow != null)
        //            {
        //                updateRow.IsCompleted = 1;
        //                db.SaveChanges();
        //            }
        //        }

        //        ret = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return ret;
        //}

        #endregion

        //Report Backup Data
        public bool TakeBackupReportData(string correctedDate)
        {
            bool result = false;
            //getting the connection string from app string.json
            string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
            string dbName = configuration.GetSection("MySettings").GetSection("Schema").Value;
            DataTable dt = new DataTable();
            string queryUnAssignedMachine = "select Distinct(Machineid) from " + dbName + ".[tblwqtyhmiscreen] where Correcteddate='" + correctedDate + "' and (Acceptreject=1 or Acceptreject1=1 ) and Isupdate=0";
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
                    var woreport = db.Tblworeport.Where(x => x.CorrectedDate == correctedDate && x.MachineId == machineId).ToList();
                    bool woReportCheck = InsertToWoReport(woreport);
                    var woLoss = db.Tblhmiscreen.Where(x => x.MachineId == machineId && x.CorrectedDate == correctedDate).Select(x => x.Hmiid).ToList();
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
                if (woLoss.Count > 0)
                {
                    foreach (int hmiid in woLoss)
                    {
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

                            db.Tblwolossess.Remove(woLossRow);
                            db.SaveChanges();

                            result = true;
                        }
                        else
                        {
                            result = true;
                        }
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

        #region Report calculation



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



        #endregion
    }
}
