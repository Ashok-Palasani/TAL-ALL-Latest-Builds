using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.EntityFrameworkCore;
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
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.DAL
{
    public class DALLoss : INoCodeInterface
    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALLoss));
        public static IConfiguration configuration;

        public DALLoss(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }
        #region
        //public Loss GetLossDet(lossmodel loss)
        //{
        //    TCF.Entity.Loss obj = new TCF.Entity.Loss();
        //    var lossdet = (from s in db.Tbllivelossofentry where (s.MachineId == loss.machineID && DateTime.Parse(s.CorrectedDate) >= DateTime.Parse(loss.FromDate) && DateTime.Parse(s.CorrectedDate) <= DateTime.Parse(loss.ToDate) && s.MessageCodeId == 999) select new { EndDateTime = s.EndDateTime,StartTime=s.StartDateTime }).ToList();
        //    foreach(var item in lossdet)
        //    {
        //        double diff = Convert.ToDateTime(item.EndDateTime).Subtract(Convert.ToDateTime(item.StartTime)).TotalMinutes;
        //            obj.MessageTypeID = "No Code";
        //        obj.startTime = (DateTime)item.StartTime;
        //        obj.endTime = (DateTime)item.EndDateTime;
        //        obj.duration = diff;
        //    }
        //    return obj;
        //}
        #endregion

        //Index Data

        #region Prv Index

        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        List<LossdetData> listLossdetData = new List<LossdetData>();
        //        string sendApprov = "";
        //        string accpetReject = "";
        //        string correctedDate = "";
        //        var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsArroval == 1).OrderBy(m => m.CorrectedDate).ToList();
        //        if (tcfLossEntry.Count > 0)
        //        {
        //            foreach (var row in tcfLossEntry)
        //            {

        //                //if (row.IsArroval == 1)
        //                //{
        //                //    sendApprov = "Sent For approval";
        //                //}
        //                //if(row.UpdateLevel == 2 && row.ApprovalLevel)
        //                if (row.IsAccept == 1)
        //                {
        //                    accpetReject = "Yes";
        //                }
        //                else if (row.IsAccept == 2)
        //                {
        //                    accpetReject = "No";
        //                }

        //                if (correctedDate != row.CorrectedDate)
        //                {
        //                    correctedDate = row.CorrectedDate;
        //                    int machineId = Convert.ToInt32(row.MachineId);
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    LossdetData objLossdetData = new LossdetData();
        //                    objLossdetData.NCID = row.Ncid;
        //                    objLossdetData.machineID = machineId;
        //                    objLossdetData.machineNmae = machineName.MachineInvNo;
        //                    objLossdetData.plantName = plantName;
        //                    objLossdetData.shopName = shopName;
        //                    objLossdetData.cellName = cellName;
        //                    objLossdetData.CorrectedDate = correctedDate;
        //                    objLossdetData.sendAppro = sendApprov;
        //                    objLossdetData.acceptReject = accpetReject;
        //                    listLossdetData.Add(objLossdetData);
        //                }
        //                //int machineId = Convert.ToInt32(row.MachineId);
        //                //string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => x.MachineDispName).FirstOrDefault();
        //                //int rid1 = Convert.ToInt32(row.ReasonLevel1);
        //                //int rid2 = Convert.ToInt32(row.ReasonLevel2);
        //                //int rid3 = Convert.ToInt32(row.ReasonLevel3);
        //                //var losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == rid1).Select(m => m.LossCode).FirstOrDefault();
        //                //var losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == rid2).Select(m => m.LossCode).FirstOrDefault();
        //                //var losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == rid3).Select(m => m.LossCode).FirstOrDefault();
        //                //LossdetData objLossdetData = new LossdetData();
        //                //objLossdetData.NCID = row.Ncid;
        //                //objLossdetData.machineID = machineId;
        //                //objLossdetData.machineNmae = machineName;
        //                //objLossdetData.startTime = Convert.ToDateTime(row.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                //objLossdetData.endTime = Convert.ToDateTime(row.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                //objLossdetData.Reasonlevel1Id = rid1;
        //                //objLossdetData.Reasonlevel1Name = losscode1;
        //                //objLossdetData.Reasonlevel2Id = rid2;
        //                //objLossdetData.Reasonlevel2Name = losscode2;
        //                //objLossdetData.Reasonlevel3Id = rid3;
        //                //objLossdetData.Reasonlevel3Name = losscode3;
        //                //objLossdetData.sendAppro = sendApprov;
        //                //objLossdetData.acceptReject = accpetReject;
        //                //listLossdetData.Add(objLossdetData);
        //            }

        //            obj.isTure = true;
        //            obj.response = listLossdetData;

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
        //    return (obj);
        //}


        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        List<LossdetData> listLossdetData = new List<LossdetData>();
        //        string firstApproval = "";
        //        string secondApproval = "";
        //        string correctedDate = "";
        //        int machineid = 0;
        //        var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsArroval == 1).OrderByDescending(m => m.MachineId).ToList();
        //        if (tcfLossEntry.Count > 0)
        //        {
        //            foreach (var row in tcfLossEntry)
        //            {
        //                if (row.IsAccept1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.IsAccept1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.IsAccept == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.IsAccept == 2)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Rejected";
        //                }

        //                if (machineid!=row.MachineId)
        //                {
        //                    //var check = listLossdetData.Find(obja.machineID=machineid);
        //                    correctedDate = row.CorrectedDate;
        //                    int machineId = Convert.ToInt32(row.MachineId);
        //                    machineid = machineId;
        //                    var machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => new { x.MachineInvNo, x.PlantId, x.ShopId, x.CellId }).FirstOrDefault();
        //                    string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == machineName.PlantId).Select(x => x.PlantName).FirstOrDefault();
        //                    string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == machineName.ShopId).Select(x => x.ShopName).FirstOrDefault();
        //                    string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == machineName.CellId).Select(x => x.CellName).FirstOrDefault();
        //                    LossdetData objLossdetData = new LossdetData();
        //                    objLossdetData.NCID = row.Ncid;
        //                    objLossdetData.machineID = machineId;
        //                    objLossdetData.machineNmae = machineName.MachineInvNo;
        //                    objLossdetData.plantName = plantName;
        //                    objLossdetData.shopName = shopName;
        //                    objLossdetData.cellName = cellName;
        //                    objLossdetData.CorrectedDate = correctedDate;
        //                    objLossdetData.firstApproval = firstApproval;
        //                    objLossdetData.secondApproval = secondApproval;
        //                    listLossdetData.Add(objLossdetData);
        //                }
        //            }

        //            obj.isTure = true;
        //            obj.response = listLossdetData;

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
        //    return (obj);
        //}
        #endregion

        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        //#region Testing only
        //        //DateTime stTime = Convert.ToDateTime("2019-09-25 00:00:00");
        //        //DateTime etTime = Convert.ToDateTime("2019-09-25 00:00:00");
        //        //DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
        //        //Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(stTime, etTime);  // for WO report updation
        //        //Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(stTime, etTime);// for OEE report updation
        //        //#endregion

        //        List<LossdetData> listLossdetData = new List<LossdetData>();

        //        string correctedDate = "";
        //        int machineID = 0;
        //        var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsArroval == 1).OrderBy(m => m.CorrectedDate).ToList();
        //        if (tcfLossEntry.Count > 0)
        //        {
        //            foreach (var row in tcfLossEntry)
        //            {
        //                string firstApproval = "";
        //                string secondApproval = "";
        //                if (row.IsAccept1 == 1)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Aprroved";

        //                }
        //                else if (row.IsAccept1 == 2)
        //                {
        //                    secondApproval = "Mail Sent and Second Level is Rejected";

        //                }
        //                if (row.IsAccept == 1)
        //                {
        //                    firstApproval = "Mail Sent and First Level is Aprroved";
        //                }
        //                else if (row.IsAccept == 2)
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
        //                    LossdetData objLossdetData = new LossdetData();
        //                    objLossdetData.NCID = row.Ncid;
        //                    objLossdetData.machineID = machineId;
        //                    objLossdetData.machineNmae = machineName.MachineInvNo;
        //                    objLossdetData.plantName = plantName;
        //                    objLossdetData.shopName = shopName;
        //                    objLossdetData.cellName = cellName;
        //                    objLossdetData.CorrectedDate = correctedDate;
        //                    objLossdetData.firstApproval = firstApproval;
        //                    objLossdetData.secondApproval = secondApproval;
        //                    listLossdetData.Add(objLossdetData);
        //                }
        //                else
        //                {
        //                    var listdata = listLossdetData.Where(m => m.machineID == row.MachineId && m.CorrectedDate == row.CorrectedDate).FirstOrDefault();
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
        //                        LossdetData objLossdetData = new LossdetData();
        //                        objLossdetData.NCID = row.Ncid;
        //                        objLossdetData.machineID = machineId;
        //                        objLossdetData.machineNmae = machineName.MachineInvNo;
        //                        objLossdetData.plantName = plantName;
        //                        objLossdetData.shopName = shopName;
        //                        objLossdetData.cellName = cellName;
        //                        objLossdetData.CorrectedDate = correctedDate;
        //                        objLossdetData.firstApproval = firstApproval;
        //                        objLossdetData.secondApproval = secondApproval;
        //                        listLossdetData.Add(objLossdetData);
        //                    }

        //                }
        //            }

        //            obj.isTure = true;
        //            obj.response = listLossdetData;

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
        //    return (obj);
        //}

        public CommonResponse Index()
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                //#region Testing only
                //DateTime stTime = Convert.ToDateTime("2019-09-25 00:00:00");
                //DateTime etTime = Convert.ToDateTime("2019-09-25 00:00:00");
                //DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
                //Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(stTime, etTime);  // for WO report updation
                //Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(stTime, etTime);// for OEE report updation
                //#endregion

                List<LossdetData> listLossdetData = new List<LossdetData>();

                string correctedDate = "";
                int machineID = 0;
                var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsArroval == 1).OrderByDescending(m => m.Ncid).ToList();
                if (tcfLossEntry.Count > 0)
                {
                    foreach (var row in tcfLossEntry)
                    {
                        string firstApproval = "";
                        string secondApproval = "";
                        if (row.IsAccept1 == 1)
                        {
                            secondApproval = "Aprroved";

                        }
                        else if (row.IsAccept1 == 2)
                        {
                            secondApproval = "Rejected";

                        }
                        if (row.IsAccept == 1)
                        {
                            firstApproval = "Aprroved";
                        }
                        else if (row.IsAccept == 2)
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
                        //    LossdetData objLossdetData = new LossdetData();
                        //    objLossdetData.NCID = row.Ncid;
                        //    objLossdetData.machineID = machineId;
                        //    objLossdetData.machineNmae = machineName.MachineInvNo;
                        //    objLossdetData.plantName = plantName;
                        //    objLossdetData.shopName = shopName;
                        //    objLossdetData.cellName = cellName;
                        //    objLossdetData.CorrectedDate = correctedDate;
                        //    objLossdetData.firstApproval = firstApproval;
                        //    objLossdetData.secondApproval = secondApproval;
                        //    listLossdetData.Add(objLossdetData);
                        //}
                        //else
                        //{
                        var listdata = listLossdetData.Where(m => m.machineID == row.MachineId && m.CorrectedDate == row.CorrectedDate).FirstOrDefault();
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
                            LossdetData objLossdetData = new LossdetData();
                            objLossdetData.NCID = row.Ncid;
                            objLossdetData.machineID = machineId;
                            objLossdetData.machineNmae = machineName.MachineInvNo;
                            objLossdetData.plantName = plantName;
                            objLossdetData.shopName = shopName;
                            objLossdetData.cellName = cellName;
                            objLossdetData.CorrectedDate = correctedDate;
                            objLossdetData.firstApproval = firstApproval;
                            objLossdetData.secondApproval = secondApproval;
                            listLossdetData.Add(objLossdetData);
                        }

                        //}
                    }

                    obj.isTure = true;
                    obj.response = listLossdetData;

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
            return (obj);
        }


        #region Prv GetLossDet

        //public EntityModelWithLossCount GetLossDet(lossmodel loss)
        //{
        //    EntityModelWithLossCount entity = new EntityModelWithLossCount();
        //    try
        //    {
        //        int plantId = loss.plantId;
        //        int shopId = loss.shopId;
        //        int cellId = loss.cellId;
        //        int machineId = loss.machineId;
        //        int lossCount = 0;


        //        //int actualSkipValue = 0;
        //        //if (loss.skipValue == 1)
        //        //{
        //        //    actualSkipValue = 0;
        //        //}
        //        //else if (loss.skipValue > 1)
        //        //{
        //        //    actualSkipValue = (loss.skipValue - 1) * loss.takeValue;
        //        //}

        //        //int actualTakeValue = loss.takeValue;
        //        var machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == loss.cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

        //        if (machineId != 0)
        //        {
        //            machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.MachineId == machineId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

        //        }
        //        else if (cellId != 0)
        //        {
        //            machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

        //        }
        //        else if (shopId != 0)
        //        {
        //            machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.ShopId == shopId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

        //        }
        //        else if (plantId != 0)
        //        {
        //            machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.PlantId == plantId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

        //        }
        //        List<Lossdet> obj1 = new List<Lossdet>();
        //        foreach (var item in machinedet)
        //        {
        //            int updateLevel = 1;
        //            var getMailIdsLevel = new List<TblTcfApprovedMaster>();

        //            getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
        //            if (getMailIdsLevel.Count == 0)
        //            {
        //                getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
        //            }
        //            foreach (var rowMail in getMailIdsLevel)
        //            {
        //                if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
        //                {
        //                    updateLevel = 2;
        //                }
        //            }



        //            var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsAccept == 0 && x.MachineId == item.machineid && x.CorrectedDate == loss.FromDate).ToList();
        //            if (tcfLossEntry.Count == 0)
        //            {

        //                //var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999 ) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).Skip(actualSkipValue).Take(actualTakeValue).ToList();

        //                var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();
        //                //var lossdet = (from s in db.Tbllivelossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();

        //                foreach (var lossitem in lossdet)
        //                {
        //                    int newTabId = 0;
        //                    var tcfloss = db.Tbltcflossofentry.Where(m => m.LossId == lossitem.Lossid).ToList();
        //                    if (tcfloss.Count == 0)
        //                    {
        //                        Tbltcflossofentry tcfobj = new Tbltcflossofentry();
        //                        tcfobj.LossId = lossitem.Lossid;
        //                        tcfobj.MachineId = item.machineid;
        //                        tcfobj.MessageCodeId = lossitem.MessageCodeId;
        //                        tcfobj.ReasonLevel1 = null;
        //                        tcfobj.ReasonLevel2 = null;
        //                        tcfobj.ReasonLevel3 = null;
        //                        tcfobj.StartDateTime = lossitem.StartTime;
        //                        tcfobj.EndDateTime = lossitem.EndTime;
        //                        tcfobj.CorrectedDate = lossitem.CorrectedDate;
        //                        tcfobj.IsUpdate = 0;
        //                        tcfobj.IsArroval = 0;
        //                        tcfobj.IsAccept = 0;
        //                        tcfobj.UpdateLevel = updateLevel;
        //                        db.Tbltcflossofentry.Add(tcfobj);
        //                        db.SaveChanges();
        //                        newTabId = tcfobj.Ncid;
        //                    }

        //                    Lossdet lossobj = new Lossdet();
        //                    //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
        //                    lossobj.NCID = newTabId;
        //                    //lossobj.MessageTypeID = "No Code";
        //                    lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                    lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                    lossobj.machineID = item.machineid;
        //                    lossobj.machineNmae = item.machinename;
        //                    //lossobj.duration = diff;
        //                    obj1.Add(lossobj);

        //                }

        //            }
        //            else if (tcfLossEntry.Count > 0)
        //            {
        //                //    bool check = false;
        //                //    foreach(var row in tcfLossEntry)
        //                //    {
        //                //        db.Tbltcflossofentry.Remove(row);
        //                //        db.SaveChanges();
        //                //        check = true;
        //                //    }
        //                //    if(check)
        //                //    {                    

        //                //        var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();
        //                //        foreach (var lossitem in lossdet)
        //                //        {
        //                //            int newTabId = 0;
        //                //            Tbltcflossofentry tcfobj = new Tbltcflossofentry();
        //                //            tcfobj.LossId = lossitem.Lossid;
        //                //            tcfobj.MachineId = item.machineid;
        //                //            tcfobj.MessageCodeId = lossitem.MessageCodeId;
        //                //            tcfobj.ReasonLevel1 = null;
        //                //            tcfobj.ReasonLevel2 = null;
        //                //            tcfobj.ReasonLevel3 = null;
        //                //            tcfobj.StartDateTime = lossitem.StartTime;
        //                //            tcfobj.EndDateTime = lossitem.EndTime;
        //                //            tcfobj.CorrectedDate = lossitem.CorrectedDate;
        //                //            tcfobj.IsUpdate = 0;
        //                //            tcfobj.IsArroval = 0;
        //                //            tcfobj.IsAccept = 0;
        //                //            tcfobj.UpdateLevel = updateLevel;
        //                //            db.Tbltcflossofentry.Add(tcfobj);
        //                //            db.SaveChanges();
        //                //            newTabId = tcfobj.Ncid;


        //                //            Lossdet lossobj = new Lossdet();
        //                //            //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
        //                //            lossobj.NCID = newTabId;
        //                //            //lossobj.MessageTypeID = "No Code";
        //                //            lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                //            lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                //            lossobj.machineID = item.machineid;
        //                //            lossobj.machineNmae = item.machinename;
        //                //            //lossobj.duration = diff;
        //                //            obj1.Add(lossobj);
        //                //        }

        //                //    }

        //                tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsAccept == 0 && x.MachineId == item.machineid && x.CorrectedDate == loss.FromDate).ToList();

        //                if (tcfLossEntry.Count > 0)
        //                {
        //                    foreach (var row in tcfLossEntry)
        //                    {
        //                        int rid1 = Convert.ToInt32(row.ReasonLevel1);
        //                        int rid2 = Convert.ToInt32(row.ReasonLevel2);
        //                        int rid3 = Convert.ToInt32(row.ReasonLevel3);
        //                        var losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == rid1).Select(m => m.LossCode).FirstOrDefault();
        //                        var losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == rid2).Select(m => m.LossCode).FirstOrDefault();
        //                        var losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == rid3).Select(m => m.LossCode).FirstOrDefault();
        //                        Lossdet lossobj = new Lossdet();
        //                        lossobj.NCID = row.Ncid;
        //                        lossobj.machineID = item.machineid;
        //                        lossobj.machineNmae = item.machinename;
        //                        lossobj.startTime = Convert.ToDateTime(row.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                        lossobj.endTime = Convert.ToDateTime(row.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss");
        //                        lossobj.Reasonlevel1Id = rid1;
        //                        lossobj.Reasonlevel1Name = losscode1;
        //                        lossobj.Reasonlevel2Id = rid2;
        //                        lossobj.Reasonlevel2Name = losscode2;
        //                        lossobj.Reasonlevel3Id = rid3;
        //                        lossobj.Reasonlevel3Name = losscode3;
        //                        obj1.Add(lossobj);
        //                    }

        //                }


        //            }
        //        }
        //        if (obj1.Count != 0)
        //        {
        //            entity.isTrue = true;
        //            entity.response = obj1;
        //            //entity.count = lossCount;
        //        }
        //        else
        //        {
        //            entity.isTrue = false;
        //            entity.response = ResourceResponse.NoItemsFound;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        entity.isTrue = false;
        //        entity.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return entity;
        //}

        #endregion

        public EntityModelWithLossCount GetLossDet(lossmodel loss)
        {
            EntityModelWithLossCount entity = new EntityModelWithLossCount();
            try
            {
                int plantId = loss.plantId;
                int shopId = loss.shopId;
                int cellId = loss.cellId;
                int machineId = loss.machineId;
                int lossCount = 0;


                var machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == loss.cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                if (machineId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.MachineId == machineId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                }
                else if (cellId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                }
                else if (shopId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.ShopId == shopId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                }
                else if (plantId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.PlantId == plantId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                }
                List<Lossdet> obj1 = new List<Lossdet>();
                foreach (var item in machinedet)
                {
                    int updateLevel = 1;
                    var getMailIdsLevel = new List<TblTcfApprovedMaster>();

                    getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
                    if (getMailIdsLevel.Count == 0)
                    {
                        getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                    }
                    foreach (var rowMail in getMailIdsLevel)
                    {
                        if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                        {
                            updateLevel = 2;
                        }
                    }



                    var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsAccept == 0 && x.MachineId == item.machineid && x.CorrectedDate == loss.FromDate).ToList();
                    if (tcfLossEntry.Count == 0)
                    {
                        var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();

                        foreach (var lossitem in lossdet)
                        {
                            int newTabId = 0;
                            bool ret = InsertIntoBackUpTable(item.machineid, loss.FromDate);
                            var tcfloss = db.Tbltcflossofentry.Where(m => m.LossId == lossitem.Lossid).ToList();
                            if (tcfloss.Count == 0)
                            {
                                Tbltcflossofentry tcfobj = new Tbltcflossofentry();
                                tcfobj.LossId = lossitem.Lossid;
                                tcfobj.MachineId = item.machineid;
                                tcfobj.MessageCodeId = lossitem.MessageCodeId;
                                tcfobj.ReasonLevel1 = null;
                                tcfobj.ReasonLevel2 = null;
                                tcfobj.ReasonLevel3 = null;
                                tcfobj.StartDateTime = lossitem.StartTime;
                                tcfobj.EndDateTime = lossitem.EndTime;
                                tcfobj.CorrectedDate = lossitem.CorrectedDate;
                                tcfobj.IsUpdate = 0;
                                tcfobj.IsArroval = 0;
                                tcfobj.IsAccept = 0;
                                tcfobj.UpdateLevel = updateLevel;
                                db.Tbltcflossofentry.Add(tcfobj);
                                db.SaveChanges();
                                newTabId = tcfobj.Ncid;
                            }

                            Lossdet lossobj = new Lossdet();
                            //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
                            lossobj.NCID = newTabId;
                            //lossobj.MessageTypeID = "No Code";
                            lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                            lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                            lossobj.machineID = item.machineid;
                            lossobj.machineNmae = item.machinename;
                            //lossobj.duration = diff;
                            obj1.Add(lossobj);

                        }

                    }
                    else if (tcfLossEntry.Count > 0)
                    {
                        //    bool check = false;
                        //    foreach(var row in tcfLossEntry)
                        //    {
                        //        db.Tbltcflossofentry.Remove(row);
                        //        db.SaveChanges();
                        //        check = true;
                        //    }
                        //    if(check)
                        //    {                    

                        //        var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();
                        //        foreach (var lossitem in lossdet)
                        //        {
                        //            int newTabId = 0;
                        //            Tbltcflossofentry tcfobj = new Tbltcflossofentry();
                        //            tcfobj.LossId = lossitem.Lossid;
                        //            tcfobj.MachineId = item.machineid;
                        //            tcfobj.MessageCodeId = lossitem.MessageCodeId;
                        //            tcfobj.ReasonLevel1 = null;
                        //            tcfobj.ReasonLevel2 = null;
                        //            tcfobj.ReasonLevel3 = null;
                        //            tcfobj.StartDateTime = lossitem.StartTime;
                        //            tcfobj.EndDateTime = lossitem.EndTime;
                        //            tcfobj.CorrectedDate = lossitem.CorrectedDate;
                        //            tcfobj.IsUpdate = 0;
                        //            tcfobj.IsArroval = 0;
                        //            tcfobj.IsAccept = 0;
                        //            tcfobj.UpdateLevel = updateLevel;
                        //            db.Tbltcflossofentry.Add(tcfobj);
                        //            db.SaveChanges();
                        //            newTabId = tcfobj.Ncid;


                        //            Lossdet lossobj = new Lossdet();
                        //            //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
                        //            lossobj.NCID = newTabId;
                        //            //lossobj.MessageTypeID = "No Code";
                        //            lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                        //            lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                        //            lossobj.machineID = item.machineid;
                        //            lossobj.machineNmae = item.machinename;
                        //            //lossobj.duration = diff;
                        //            obj1.Add(lossobj);
                        //        }

                        //    }

                        tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsAccept == 0 && x.MachineId == item.machineid && x.CorrectedDate == loss.FromDate).ToList();

                        if (tcfLossEntry.Count > 0)
                        {
                            foreach (var row in tcfLossEntry)
                            {
                                int rid1 = Convert.ToInt32(row.ReasonLevel1);
                                int rid2 = Convert.ToInt32(row.ReasonLevel2);
                                int rid3 = Convert.ToInt32(row.ReasonLevel3);
                                var losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == rid1).Select(m => m.LossCode).FirstOrDefault();
                                var losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == rid2).Select(m => m.LossCode).FirstOrDefault();
                                var losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == rid3).Select(m => m.LossCode).FirstOrDefault();
                                Lossdet lossobj = new Lossdet();
                                lossobj.NCID = row.Ncid;
                                lossobj.machineID = item.machineid;
                                lossobj.machineNmae = item.machinename;
                                lossobj.startTime = Convert.ToDateTime(row.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                                lossobj.endTime = Convert.ToDateTime(row.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                                lossobj.Reasonlevel1Id = rid1;
                                lossobj.Reasonlevel1Name = losscode1;
                                lossobj.Reasonlevel2Id = rid2;
                                lossobj.Reasonlevel2Name = losscode2;
                                lossobj.Reasonlevel3Id = rid3;
                                lossobj.Reasonlevel3Name = losscode3;
                                obj1.Add(lossobj);
                            }
                        }
                    }
                }
                if (obj1.Count != 0)
                {
                    entity.isTrue = true;
                    entity.response = obj1.OrderBy(x => x.startTime);
                    //entity.count = lossCount;
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }
        
        public bool InsertIntoBackUpTable(int machineId, string correctedDate)
        {
            bool ret = false;
            try
            {

                var tcflossBackup = db.TblBachUplossofentry.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate).ToList();
                if (tcflossBackup.Count == 0)
                {
                    var lossdet = (from s in db.Tbllossofentry where (s.MachineId == machineId && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(correctedDate)) && s.MessageCodeId == 999) select s).ToList();
                    foreach (var data in lossdet)
                    {
                        TblBachUplossofentry obj = new TblBachUplossofentry();
                        obj.CorrectedDate = data.CorrectedDate;
                        obj.DoneWithRow = data.DoneWithRow;
                        obj.EndDateTime = data.EndDateTime;
                        obj.EntryTime = data.EntryTime;
                        obj.ForRefresh = data.ForRefresh;
                        obj.IsScreen = data.IsScreen;
                        obj.IsStart = data.IsStart;
                        obj.IsUpdate = data.IsUpdate;
                        obj.LossMonth = data.LossMonth;
                        obj.LossQuarter = data.LossQuarter;
                        obj.LossWeekNumber = data.LossWeekNumber;
                        obj.LossYear = data.LossYear;
                        obj.MachineId = data.MachineId;
                        obj.MessageCode = data.MessageCode;
                        obj.MessageCodeId = data.MessageCodeId;
                        obj.MessageDesc = data.MessageDesc;
                        obj.Shift = data.Shift;
                        obj.StartDateTime = data.StartDateTime;
                        db.TblBachUplossofentry.Add(obj);
                        db.SaveChanges();
                        ret = true;
                    }
                }
                else
                {
                    ret = false;
                }

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        public CommonResponse DeletePreviousNoCode(lossmodel loss)
        {
            CommonResponse entity = new CommonResponse();
            try
            {

                int plantId = loss.plantId;
                int shopId = loss.shopId;
                int cellId = loss.cellId;
                int machineId = loss.machineId;
                var machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == loss.cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();

                if (machineId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.MachineId == machineId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();
                }
                else if (cellId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.CellId == cellId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();
                }
                else if (shopId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.ShopId == shopId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();
                }
                else if (plantId != 0)
                {
                    machinedet = (from s in db.Tblmachinedetails where (s.IsDeleted == 0 && s.PlantId == plantId) select new { machineid = s.MachineId, machinename = s.MachineInvNo }).ToList();
                }
                List<Lossdet> obj1 = new List<Lossdet>();
                foreach (var item in machinedet)
                {
                    int updateLevel = 1;
                    var getMailIdsLevel = new List<TblTcfApprovedMaster>();

                    getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.CellId == cellId).ToList();
                    if (getMailIdsLevel.Count == 0)
                    {
                        getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 1 && x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                    }
                    foreach (var rowMail in getMailIdsLevel)
                    {
                        if (rowMail.SecondApproverCcList != "" && rowMail.SecondApproverToList != "")
                        {
                            updateLevel = 2;
                        }
                    }

                    var tcfLossEntry = db.Tbltcflossofentry.Where(x => x.IsArroval == 1 && x.MachineId == loss.machineId && x.CorrectedDate == loss.FromDate).OrderByDescending(m => m.Ncid).ToList();
                    if (tcfLossEntry.Count > 0)
                    {
                        foreach (var lossrow in tcfLossEntry)
                        {
                            if (lossrow.IsAccept == 1)
                            {
                                entity.isTure = false;
                                entity.errorMsg = "For this Machine,Mail Sent, and First level is Approved";
                                break;
                            }
                            else if (lossrow.IsArroval == 1)
                            {
                                entity.isTure = false;
                                entity.errorMsg = "For this Machine,Mail Sent, and Approval is pending";
                                break;
                            }
                        }
                    }
                    #region Not Required
                    //if (tcfLossEntry.Count > 0)
                    //{
                    //    var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();
                    //    //var lossdet = (from s in db.Tbllivelossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();

                    //    foreach (var lossitem in lossdet)
                    //    {
                    //        int newTabId = 0;
                    //        var tcfloss = db.Tbltcflossofentry.Where(m => m.LossId == lossitem.Lossid).ToList();
                    //        if (tcfloss.Count == 0)
                    //        {
                    //            Tbltcflossofentry tcfobj = new Tbltcflossofentry();
                    //            tcfobj.LossId = lossitem.Lossid;
                    //            tcfobj.MachineId = item.machineid;
                    //            tcfobj.MessageCodeId = lossitem.MessageCodeId;
                    //            tcfobj.ReasonLevel1 = null;
                    //            tcfobj.ReasonLevel2 = null;
                    //            tcfobj.ReasonLevel3 = null;
                    //            tcfobj.StartDateTime = lossitem.StartTime;
                    //            tcfobj.EndDateTime = lossitem.EndTime;
                    //            tcfobj.CorrectedDate = lossitem.CorrectedDate;
                    //            tcfobj.IsUpdate = 0;
                    //            tcfobj.IsArroval = 0;
                    //            tcfobj.IsAccept = 0;
                    //            tcfobj.UpdateLevel = updateLevel;
                    //            db.Tbltcflossofentry.Add(tcfobj);
                    //            db.SaveChanges();
                    //            newTabId = tcfobj.Ncid;
                    //        }

                    //        Lossdet lossobj = new Lossdet();
                    //        //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
                    //        lossobj.NCID = newTabId;
                    //        //lossobj.MessageTypeID = "No Code";
                    //        lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                    //        lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                    //        lossobj.machineID = item.machineid;
                    //        lossobj.machineNmae = item.machinename;
                    //        //lossobj.duration = diff;
                    //        obj1.Add(lossobj);
                    //    }
                    //}
                    #endregion
                   else if (tcfLossEntry.Count == 0)
                    {
                        bool check = false;
                        foreach (var row in tcfLossEntry)
                        {
                            db.Tbltcflossofentry.Remove(row);
                            db.SaveChanges();
                            check = true;
                        }
                        if (check)
                        {

                            var lossdet = (from s in db.Tbllossofentry where (s.MachineId == item.machineid && (DateTime.Parse(s.CorrectedDate) == DateTime.Parse(loss.FromDate)) && s.MessageCodeId == 999) select new { EndTime = s.EndDateTime, StartTime = s.StartDateTime, Lossid = s.LossId, s.MessageCodeId, CorrectedDate = s.CorrectedDate }).ToList();
                            foreach (var lossitem in lossdet)
                            {
                                int newTabId = 0;
                                Tbltcflossofentry tcfobj = new Tbltcflossofentry();
                                tcfobj.LossId = lossitem.Lossid;
                                tcfobj.MachineId = item.machineid;
                                tcfobj.MessageCodeId = lossitem.MessageCodeId;
                                tcfobj.ReasonLevel1 = null;
                                tcfobj.ReasonLevel2 = null;
                                tcfobj.ReasonLevel3 = null;
                                tcfobj.StartDateTime = lossitem.StartTime;
                                tcfobj.EndDateTime = lossitem.EndTime;
                                tcfobj.CorrectedDate = lossitem.CorrectedDate;
                                tcfobj.IsUpdate = 0;
                                tcfobj.IsArroval = 0;
                                tcfobj.IsAccept = 0;
                                tcfobj.UpdateLevel = updateLevel;
                                db.Tbltcflossofentry.Add(tcfobj);
                                db.SaveChanges();
                                newTabId = tcfobj.Ncid;


                                Lossdet lossobj = new Lossdet();
                                //double diff = Convert.ToDateTime(lossitem.EndTime).Subtract(Convert.ToDateTime(lossitem.StartTime)).TotalMinutes;
                                lossobj.NCID = newTabId;
                                //lossobj.MessageTypeID = "No Code";
                                lossobj.startTime = Convert.ToDateTime(lossitem.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                                lossobj.endTime = Convert.ToDateTime(lossitem.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                                lossobj.machineID = item.machineid;
                                lossobj.machineNmae = item.machinename;
                                //lossobj.duration = diff;
                                obj1.Add(lossobj);
                            }
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


        #region Split Duration 

        //Split duration
        public CommonResponse SplitDuration(int uawoid)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var unAssignedWodet = db.Tbltcflossofentry.Where(x => x.Ncid == uawoid).FirstOrDefault();
                if (unAssignedWodet != null)
                {
                    string endTime = Convert.ToDateTime(unAssignedWodet.EndDateTime).ToString("HH:mm:ss");
                    string[] splitEndTime = endTime.Split(':');
                    SplitDurationList splitDurationList = new SplitDurationList();
                    List<SplitDuration> splitDurations = new List<SplitDuration>();
                    SplitDuration objSplitDuration = new SplitDuration();
                    objSplitDuration.startTime = Convert.ToDateTime(unAssignedWodet.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                    objSplitDuration.endTime = Convert.ToDateTime(unAssignedWodet.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                    //objSplitDuration.endHour = splitEndTime[0];
                    //objSplitDuration.endMinute = splitEndTime[1];
                    //objSplitDuration.endSecond = splitEndTime[2];
                    objSplitDuration.uaWOId = Convert.ToInt32(unAssignedWodet.Ncid);
                    splitDurations.Add(objSplitDuration);
                    splitDurationList.uaWOIds = unAssignedWodet.Ncid.ToString();
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
                var uaWoData = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).FirstOrDefault();
                if (uaWoData != null)
                {
                    DateTime endDateTime = Convert.ToDateTime(data.endTime);
                    DateTime endDateTimePrv = Convert.ToDateTime(endDateTime.AddDays(-1));
                    DateTime endTime = Convert.ToDateTime(uaWoData.EndDateTime);
                    DateTime startTime = Convert.ToDateTime(uaWoData.StartDateTime);
                    int durationCheckFirst = 0;
                    int durationCheckSecond = 0;
                    if (endDateTime <= endTime && endDateTime >= startTime)
                    {
                        durationCheckFirst = Convert.ToInt32(endDateTime.Subtract(startTime).TotalSeconds);
                        durationCheckSecond = Convert.ToInt32(endTime.Subtract(endDateTime).TotalSeconds);
                        if (durationCheckFirst > 120 && durationCheckSecond > 120)
                        {
                            bool check = ValidatePrvEndTime(allIds, endDateTime);
                            if (check)
                            {
                                obj.isTure = false;
                                obj.response = endTime + " This Time Already Exist";
                            }
                            else
                            {
                                string nowDate = uaWoData.CorrectedDate;
                                uaWoData.EndDateTime = endDateTime;
                                db.SaveChanges();
                                Tbltcflossofentry addRow = new Tbltcflossofentry();
                                addRow.StartDateTime = endDateTime;
                                addRow.EndDateTime = endTime;
                                addRow.LossId = uaWoData.LossId;
                                addRow.MessageCodeId = 999;
                                //addRow.Isdeleted = 0;
                                //addRow.Createdby = 1;// change as per user login
                                //addRow.Createdon = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                addRow.CorrectedDate = nowDate;
                                addRow.MachineId = uaWoData.MachineId;
                                addRow.IsSplitDuration = uaWoId; // need to add column in table
                                db.Tbltcflossofentry.Add(addRow);
                                db.SaveChanges();
                                allIds += "," + addRow.Ncid;
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
                            obj.response = "The Duration splited must be greater than 120 Seconds";
                        }
                    }
                    else if (endDateTimePrv <= endTime && endDateTimePrv >= startTime)
                    {
                        durationCheckFirst = Convert.ToInt32(endDateTimePrv.Subtract(startTime).TotalSeconds);
                        durationCheckSecond = Convert.ToInt32(endTime.Subtract(endDateTimePrv).TotalSeconds);

                        if (durationCheckFirst > 120 && durationCheckSecond > 120)
                        {

                            bool check = ValidatePrvEndTime(allIds, endDateTimePrv);
                            if (check)
                            {
                                obj.isTure = false;
                                obj.response = endTime + " This Time Already Exist";
                            }
                            else
                            {
                                string nowDate = uaWoData.CorrectedDate;
                                uaWoData.EndDateTime = endDateTimePrv;
                                db.SaveChanges();
                                Tbltcflossofentry addRow = new Tbltcflossofentry();
                                addRow.StartDateTime = endDateTimePrv;
                                addRow.EndDateTime = endTime;
                                addRow.MessageCodeId = 999;
                                //addRow.Isdeleted = 0;
                                //addRow.Createdby = 1;// change as per user login
                                //addRow.Createdon = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                addRow.CorrectedDate = nowDate;
                                addRow.MachineId = uaWoData.MachineId;
                                addRow.IsSplitDuration = uaWoId;// column need to be added
                                db.Tbltcflossofentry.Add(addRow);
                                db.SaveChanges();
                                allIds += "," + addRow.Ncid;
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
                            obj.response = "The Duration splited must be greater than 120 Seconds";
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
                    var check = db.Tbltcflossofentry.Where(x => x.Ncid == id && x.EndDateTime == endDateTime).FirstOrDefault();
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
                var uaWoData = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).FirstOrDefault();
                objSplitDuration.startTime = Convert.ToDateTime(uaWoData.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                objSplitDuration.endTime = Convert.ToDateTime(uaWoData.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                objSplitDuration.uaWOId = uaWoData.Ncid;
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
                            var unAssignedData = db.Tbltcflossofentry.Where(x => x.Ncid == spliId).FirstOrDefault();
                            int updateId = Convert.ToInt32(uaWOIDSList[i - 1]);
                            var updateRow = db.Tbltcflossofentry.Where(x => x.Ncid == updateId).FirstOrDefault();
                            updateRow.EndDateTime = unAssignedData.EndDateTime;
                            db.SaveChanges();
                            db.Remove(unAssignedData);
                            db.SaveChanges();
                            //obj.isTure = true;
                            //obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                        else
                        {
                            var unAssignedData = db.Tbltcflossofentry.Where(x => x.Ncid == spliId).FirstOrDefault();
                            int updateId = Convert.ToInt32(uaWOIDSList[i + 1]);
                            var updateRow = db.Tbltcflossofentry.Where(x => x.Ncid == updateId).FirstOrDefault();
                            updateRow.StartDateTime = unAssignedData.StartDateTime;
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
                DateTime endDateTime = Convert.ToDateTime(db.Tbltcflossofentry.Where(x => x.Ncid == uaWoIdLast).Select(x => x.EndDateTime).FirstOrDefault());
                for (int i = 1; i < count; i++)
                {
                    uaWoIdLast = Convert.ToInt32(splitIds[i]);
                    var unawoRemoveRow = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoIdLast).FirstOrDefault();
                    db.Tbltcflossofentry.Remove(unawoRemoveRow);
                    db.SaveChanges();
                }
                var unawoUpdateRow = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).FirstOrDefault();
                unawoUpdateRow.EndDateTime = endDateTime;
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
                    int uaWoId = Convert.ToInt32(splitIds[count]);
                    DateTime endDateTime = Convert.ToDateTime(db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).Select(x => x.EndDateTime).FirstOrDefault());
                    for (int i = 1; i < count; i++)
                    {
                        uaWoId = Convert.ToInt32(splitIds[i]);
                        var unawoRemoveRow = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).FirstOrDefault();
                        db.Tbltcflossofentry.Remove(unawoRemoveRow);
                        db.SaveChanges();
                    }
                    var unawoUpdateRow = db.Tbltcflossofentry.Where(x => x.Ncid == uaWoId).FirstOrDefault();
                    unawoUpdateRow.EndDateTime = endDateTime;
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

        #endregion



        #region PlantShopCellDetails


        //Get plant details
        //public EntityModel GetPlantDet()
        //{
        //    EntityModel entiy = new EntityModel();
        //    List<Plant> plantlist = new List<Plant>();
        //    var plantdet = db.Tblplant.Where(m => m.IsDeleted == 0).ToList();
        //    foreach (var item in plantdet)
        //    {
        //        Plant plantobj = new Plant();
        //        plantobj.plantid = item.PlantId;
        //        plantobj.plantdesc = item.PlantDesc;
        //        plantlist.Add(plantobj);
        //    }
        //    if (plantlist != null)
        //    {
        //        entiy.isTrue = true;
        //        entiy.response = plantlist;
        //    }
        //    else
        //    {
        //        entiy.isTrue = false;
        //        entiy.response = "fail";
        //    }
        //    return entiy;
        //}

        //Get Shop details
        //public EntityModel GetShopDet(int plantid)
        //{
        //    EntityModel entiy = new EntityModel();
        //    List<Shop> shoplist = new List<Shop>();
        //    var shopdet = db.Tblshop.Where(m => m.IsDeleted == 0 && m.PlantId == plantid).ToList();
        //    foreach (var item in shopdet)
        //    {
        //        Shop shopobj = new Shop();
        //        shopobj.Shopid = item.ShopId;
        //        shopobj.Shopdesc = item.ShopDesc;
        //        shoplist.Add(shopobj);
        //    }
        //    if (shoplist != null)
        //    {
        //        entiy.isTrue = true;
        //        entiy.response = shoplist;
        //    }
        //    else
        //    {
        //        entiy.isTrue = false;
        //        entiy.response = "fail";
        //    }
        //    return entiy;
        //}

        ////Get Cell details
        //public EntityModel GetCellDet(int shopid)
        //{
        //    EntityModel entiy = new EntityModel();

        //    List<Cell> celllist = new List<Cell>();
        //    var celldet = db.Tblcell.Where(m => m.IsDeleted == 0 && m.ShopId == shopid).ToList();
        //    foreach (var item in celldet)
        //    {
        //        Cell cellobj = new Cell();
        //        cellobj.Cellid = item.CellId;
        //        cellobj.Celldesc = item.CellDesc;
        //        celllist.Add(cellobj);
        //    }
        //    if (celllist != null)
        //    {
        //        entiy.isTrue = true;
        //        entiy.response = celllist;
        //    }
        //    else
        //    {
        //        entiy.isTrue = false;
        //        entiy.response = "fail";
        //    }
        //    return entiy;
        //}

        #endregion


        //Get LossCodeLevel 1
        public EntityModel GetLossCodeLevel1()
        {
            EntityModel entity = new EntityModel();
            try
            {
                List<LossCodeLevel1> losslist = new List<LossCodeLevel1>();
                var lossdet = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodesLevel == 1 && m.LossCode != "999" && m.LossCodeDesc != "Breakdown").ToList();
                foreach (var row in lossdet)
                {
                    LossCodeLevel1 lossobj1 = new LossCodeLevel1();
                    lossobj1.LossCodeID = row.LossCodeId;
                    lossobj1.LossCode = row.LossCode;
                    losslist.Add(lossobj1);
                }
                //if (losslist != null)
                if (losslist.Count > 0)
                {
                    entity.isTrue = true;
                    entity.response = losslist;
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        //Get LossCodeLevel 2
        public EntityModel GetLossCodeLevel2(int lossCodeID)
        {
            EntityModel entity = new EntityModel();
            try
            {
                List<LossCodeLevel2> losscodelist = new List<LossCodeLevel2>();
                var lossdet = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodesLevel1Id == lossCodeID && m.LossCodesLevel == 2).ToList();
                foreach (var row in lossdet)
                {
                    LossCodeLevel2 lossobj = new LossCodeLevel2();
                    lossobj.LossCode = row.LossCode;
                    lossobj.LossCodeID = row.LossCodeId;
                    losscodelist.Add(lossobj);
                }
                //if (losscodelist != null)
                if (losscodelist.Count > 0)
                {
                    entity.isTrue = true;
                    entity.response = losscodelist;
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        //Get LossCodeLevel 3
        public EntityModel GetLossCodeLevel3(int lossCodeID)
        {
            EntityModel entity = new EntityModel();
            try
            {
                List<LossCodeLevel3> losslist = new List<LossCodeLevel3>();
                var lossdet = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodesLevel2Id == lossCodeID && m.LossCodesLevel == 3).ToList();
                foreach (var row in lossdet)
                {
                    LossCodeLevel3 lossobj = new LossCodeLevel3();
                    lossobj.LossCode = row.LossCode;
                    lossobj.LossCodeID = row.LossCodeId;
                    losslist.Add(lossobj);
                }
                //if (losslist!= null)
                if (losslist.Count > 0)
                {
                    entity.isTrue = true;
                    entity.response = losslist;
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;

        }

        //Update tcf table when click on set button
        public EntityModel UpdatetcfLoss(updateLoss tcfdata)
        {
            string losscode = "";
            int losscodeid = 0;
            EntityModel entity = new EntityModel();
            try
            {
                var tcfrecord = db.Tbltcflossofentry.Where(m => m.Ncid == tcfdata.NCID).FirstOrDefault();
                if (tcfrecord != null)
                {
                    if (tcfdata.reasonLevel3Id != 0)
                    {
                        losscodeid = tcfdata.reasonLevel3Id;
                        tcfrecord.NoOfReason = tcfdata.NoOfReason;
                        tcfrecord.MessageCodeId = losscodeid;
                        tcfrecord.ReasonLevel1 = Convert.ToString(tcfdata.reasonLevel1Id);
                        tcfrecord.ReasonLevel2 = Convert.ToString(tcfdata.reasonLevel2Id);
                        tcfrecord.ReasonLevel3 = Convert.ToString(tcfdata.reasonLevel3Id);
                        //losscode = db.Tbllossescodes.Where(m => m.LossCodeId == tcfdata.reasonLevel3Id).Select(m => m.LossCode).FirstOrDefault();
                    }
                    else if (tcfdata.reasonLevel2Id != 0)
                    {
                        losscodeid = tcfdata.reasonLevel2Id;
                        tcfrecord.NoOfReason = tcfdata.NoOfReason;
                        tcfrecord.MessageCodeId = losscodeid;
                        tcfrecord.ReasonLevel1 = Convert.ToString(tcfdata.reasonLevel1Id);
                        tcfrecord.ReasonLevel2 = Convert.ToString(tcfdata.reasonLevel2Id);
                        tcfrecord.ReasonLevel3 = null;
                        //losscode = db.Tbllossescodes.Where(m => m.LossCodeId == tcfdata.reasonLevel2Id).Select(m => m.LossCode).FirstOrDefault();
                    }
                    else if (tcfdata.reasonLevel1Id != 0)
                    {
                        losscodeid = tcfdata.reasonLevel1Id;
                        tcfrecord.NoOfReason = tcfdata.NoOfReason;
                        tcfrecord.MessageCodeId = losscodeid;
                        tcfrecord.ReasonLevel1 = Convert.ToString(tcfdata.reasonLevel1Id);
                        tcfrecord.ReasonLevel2 = null;
                        tcfrecord.ReasonLevel3 = null;
                        //losscode = db.Tbllossescodes.Where(m => m.LossCodeId == tcfdata.reasonLevel1Id).Select(m => m.LossCode).FirstOrDefault();
                    }


                    //tcfrecord.IsArroval = 1;
                    db.Entry(tcfrecord).State = EntityState.Modified;
                    db.SaveChanges();
                    entity.isTrue = true;
                    entity.response = ResourceResponse.SuccessMessage;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        //when click on Set to Approval button Get tcfLoss details
        //public EntityModel GettcfLoss()
        //{
        //    EntityModel enity = new EntityModel();
        //    List<reasonlevel> reasonlevellist = new List<reasonlevel>();
        //    var tcfloss = db.Tbltcflossofentry.Where(m => m.IsArroval == 1 && m.IsAccept==0).ToList();
        //    foreach(var item in tcfloss)
        //    {
        //        reasonlevel obj = new reasonlevel();
        //        obj.reasonLevel1 = item.ReasonLevel1;
        //        obj.reasonLevel2 = item.ReasonLevel2;
        //        obj.reasonLevel3 = item.ReasonLevel3;
        //        obj.StartTime = item.StartDateTime;
        //        obj.EndTime = item.EndDateTime;
        //        reasonlevellist.Add(obj);
        //    }
        //    if (reasonlevellist != null)
        //    {
        //        enity.isTrue = true;
        //        enity.response = reasonlevellist;
        //    }
        //    else
        //    {
        //        enity.isTrue = false;
        //        enity.response = "fail";
        //    }
        //    try
        //    {
        //       // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");
        //        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\ReasonEmailTemplate.html");
        //        string htmlStr = File.ReadAllText(reader);

        //        String[] seperator = { "{ {reasonStart}}" };
        //        string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

        //        var reasonHtml = htmlArr[1].Split(new String[] { "{{reasonEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
        //        htmlStr = htmlStr.Replace("{{reasonStart}}", "");
        //        htmlStr = htmlStr.Replace("{{reasonEnd}}", "");
        //        for (int i = 0; i < reasonlevellist.Count; i++)
        //        {
        //            String Stime = Convert.ToString(reasonlevellist[i].StartTime);
        //            String Etime = Convert.ToString(reasonlevellist[i].EndTime);
        //            htmlStr = htmlStr.Replace("{{rl1}}", reasonlevellist[i].reasonLevel1);
        //            htmlStr = htmlStr.Replace("{{rl2}}", reasonlevellist[i].reasonLevel2);
        //            htmlStr = htmlStr.Replace("{{rl3}}", reasonlevellist[i].reasonLevel3);
        //            htmlStr = htmlStr.Replace("{{startTime}}", Stime);
        //            htmlStr = htmlStr.Replace("{{endTime}}", Etime);
        //            if (reasonlevellist.Count == 1)
        //            {
        //                htmlStr = htmlStr.Replace("{{reason}}", "");
        //            }
        //            else if (i < reasonlevellist.Count - 1)
        //            {

        //               htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
        //            }
        //        }

        //        htmlStr = htmlStr.Replace("{{reason}}", "");
        //        htmlStr = htmlStr.Replace("{{userName}}", "Vignesh");
        //        htmlStr = htmlStr.Replace("{{SupervisorName}}", "Monika");

        //        string toMailID = "vignesh.pai@srkssolutions.com";
        //        string ccMailID = "monika.ms@srkssolutions.com";
        //        MailMessage mail = new MailMessage();
        //        mail.To.Add(toMailID);
        //        mail.CC.Add(ccMailID);
        //        mail.From = new MailAddress("monika.ms@srkssolutions.com");
        //        mail.Subject = "test mail";
        //        mail.Body = ""+ htmlStr;
        //        mail.IsBodyHtml = true;
        //        SmtpClient smtp = new SmtpClient();
        //        smtp.Host = "smtp.gmail.com";
        //        smtp.Port = 587;
        //        smtp.EnableSsl = true;
        //        smtp.UseDefaultCredentials = false;
        //        smtp.Credentials = new System.Net.NetworkCredential("monika.ms@srkssolutions.com", "monika.ms10$");
        //        smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
        //        smtp.Send(mail);
        //    }
        //    catch(Exception ex)
        //    {

        //    }
        //    return enity;
        //}


        #region
        //when click on Enter Reason button Get tcfLoss details  


        //public EntityModel GettcfLossLevel(int LossID)
        //{
        //    EntityModel enity = new EntityModel();
        //    try
        //    {
        //        reason obj = new reason();
        //        List<reason> reasonlist = new List<reason>();
        //        var tcfloss = db.Tbltcflossofentry.Where(m => m.IsUpdate == 0 && m.IsArroval == 0 && m.IsAccept == 0 && m.LossId == LossID).ToList();
        //        foreach (var row in tcfloss)
        //        {
        //            obj.StartTime = row.StartDateTime;
        //            obj.EndTime = row.EndDateTime;
        //            reasonlist.Add(obj);
        //        }
        //        if (reasonlist != null)
        //        {
        //            enity.isTrue = true;
        //            enity.response = reasonlist;
        //        }
        //        else
        //        {
        //            enity.isTrue = false;
        //            enity.response = ResourceResponse.FailureMessage;
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        enity.isTrue = false;
        //        enity.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return enity;
        //}
        #endregion


        //when click on Accept button update all the records to tbllivelossofentry  
        #region Old Live loss
        //public EntityModel UpdateliveLoss(string correctedDate)
        //{
        //    int losscodeid = 0;
        //    string losscode = "";
        //    EntityModel entity = new EntityModel();

        //    var tcfrecord = db.Tbltcflossofentry.Where(m => (m.IsAccept == 1 || m.IsAccept1 == 1) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate).ToList();
        //    foreach (var row in tcfrecord)
        //    {
        //        try
        //        {

        //            if (row.ReasonLevel3 != null)
        //            {
        //                losscodeid = Convert.ToInt32(row.ReasonLevel3);
        //                losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
        //            }
        //            else if (row.ReasonLevel2 != null)
        //            {
        //                losscodeid = Convert.ToInt32(row.ReasonLevel2);
        //                losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
        //            }
        //            else if (row.ReasonLevel1 != null)
        //            {
        //                losscodeid = Convert.ToInt32(row.ReasonLevel1);
        //                losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
        //            }

        //            var livelossdet = db.Tbllivelossofentry.Where(m => m.LossId == row.LossId).FirstOrDefault();
        //            var histLossDetails = db.Tbllossofentry.Where(m => m.LossId == row.LossId).FirstOrDefault();
        //            if (livelossdet != null)
        //            {
        //                livelossdet.MessageCode = losscode;
        //                livelossdet.MessageCodeId = losscodeid;
        //                livelossdet.MessageDesc = losscode;
        //                db.Entry(livelossdet).State = EntityState.Modified;
        //                db.SaveChanges();
        //            }
        //            if (histLossDetails != null)
        //            {
        //                histLossDetails.MessageCode = losscode;
        //                histLossDetails.MessageCodeId = losscodeid;
        //                histLossDetails.MessageDesc = losscode;
        //                db.Entry(histLossDetails).State = EntityState.Modified;
        //                db.SaveChanges();
        //            }

        //            row.IsUpdate = 1;
        //            db.Entry(row).State = EntityState.Modified;
        //            db.SaveChanges();


        //            //DateTime stTime = Convert.ToDateTime(tcfdata.CorrectedDate + " " + "00:00:00");
        //            //DateTime etTime = Convert.ToDateTime(tcfdata.CorrectedDate + " " + "00:00:00");
        //            DateTime stTime = Convert.ToDateTime(correctedDate);
        //            DateTime etTime = Convert.ToDateTime(correctedDate);
        //            Tbloeedashboardvariables oeedet = db.Tbloeedashboardvariables.Where(m => m.Wcid == row.MachineId && m.StartDate >= stTime.Date && m.EndDate <= etTime.Date).FirstOrDefault();

        //            TblBackupoeedashboardvariables objback = new TblBackupoeedashboardvariables();
        //            objback.Blue = oeedet.Blue;
        //            objback.OldOeeId = oeedet.OeevariablesId;
        //            objback.CellId = oeedet.CellId;
        //            objback.CreatedBy = oeedet.CreatedBy;
        //            objback.CreatedOn = oeedet.CreatedOn;
        //            objback.DownTimeBreakdown = oeedet.DownTimeBreakdown;
        //            objback.EndDate = oeedet.EndDate;
        //            objback.Green = oeedet.Green;
        //            objback.IsDeleted = oeedet.IsDeleted;
        //            objback.Loss1Name = oeedet.Loss1Name;
        //            objback.Loss1Value = oeedet.Loss1Value;
        //            objback.Loss2Name = oeedet.Loss2Name;
        //            objback.Loss2Value = oeedet.Loss2Value;
        //            objback.Loss3Name = oeedet.Loss3Name;
        //            objback.Loss3Value = oeedet.Loss3Value;
        //            objback.Loss4Name = oeedet.Loss4Name;
        //            objback.Loss4Value = oeedet.Loss4Value;
        //            objback.Loss5Name = oeedet.Loss5Name;
        //            objback.Loss5Value = oeedet.Loss5Value;
        //            objback.MinorLosses = oeedet.MinorLosses;
        //            objback.PlantId = oeedet.PlantId;
        //            objback.ReWotime = oeedet.ReWotime;
        //            objback.Roalossess = oeedet.Roalossess;
        //            objback.ScrapQtyTime = oeedet.ScrapQtyTime;
        //            objback.SettingTime = oeedet.SettingTime;
        //            objback.ShopId = oeedet.ShopId;
        //            objback.StartDate = oeedet.StartDate;
        //            objback.SummationOfSctvsPp = oeedet.SummationOfSctvsPp;
        //            objback.Wcid = oeedet.Wcid;
        //            db.TblBackupoeedashboardvariables.Add(objback);
        //            db.SaveChanges();

        //            db.Tbloeedashboardvariables.Remove(oeedet);
        //            db.SaveChanges();

        //            TakeBackupReportData(correctedDate);

        //            log.Error("Before Method");
        //            //bool calres = CalculateOEEForYesterday(stTime, etTime);
        //            Task<bool> reportWOUpdate = CalWODataForYesterday(stTime, etTime);  // for WO report updation
        //            Task<bool> reportOEEUpdate = CalculateOEEForYesterday(stTime, etTime);// for OEE report updation

        //            if (reportOEEUpdate.Result == true && reportOEEUpdate.Result == true)
        //            {
        //                entity.isTrue = true;
        //                entity.response = ResourceResponse.SuccessMessage;
        //            }

        //            //if (calres == true)
        //            //{
        //            //    entity.isTrue = true;
        //            //    log.Error("CalculateOEEForYesterday Succeed");
        //            //    entity.response = ResourceResponse.SuccessMessage;
        //            //}
        //            //else
        //            //{
        //            //    entity.isTrue = false;
        //            //    log.Error("CalculateOEEForYesterday failed");
        //            //    entity.response = ResourceResponse.ExceptionMessage;
        //            //}

        //        }
        //        catch (Exception ex)
        //        {
        //            entity.isTrue = false;
        //            entity.response = ResourceResponse.ExceptionMessage;
        //            log.Error("exception");
        //            log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        }
        //    }

        //    return entity;
        //}
        #endregion

        public EntityModel UpdateliveLoss(string correctedDate,int MachineID)
        {
            int losscodeid = 0;
            string losscode = "";
            bool check = false;
            EntityModel entity = new EntityModel();
            try
            {
                CheckSplitDurationLossUpdation(correctedDate);
                var tcfrecord = db.Tbltcflossofentry.Where(m => (m.IsAccept == 1 || m.IsAccept1 == 1) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate).ToList();
                foreach (var row in tcfrecord)
                {
                    if (row.ReasonLevel3 != null)
                    {
                        losscodeid = Convert.ToInt32(row.ReasonLevel3);
                        losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
                    }
                    else if (row.ReasonLevel2 != null)
                    {
                        losscodeid = Convert.ToInt32(row.ReasonLevel2);
                        losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
                    }
                    else if (row.ReasonLevel1 != null)
                    {
                        losscodeid = Convert.ToInt32(row.ReasonLevel1);
                        losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).Select(m => m.LossCode).FirstOrDefault();
                    }

                    var livelossdet = db.Tbllivelossofentry.Where(m => m.LossId == row.LossId).FirstOrDefault();
                    var histLossDetails = db.Tbllossofentry.Where(m => m.LossId == row.LossId).FirstOrDefault();
                    if (livelossdet != null)
                    {
                        livelossdet.MessageCode = losscode;
                        livelossdet.MessageCodeId = losscodeid;
                        livelossdet.MessageDesc = losscode;
                        db.Entry(livelossdet).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    if (histLossDetails != null)
                    {
                        histLossDetails.MessageCode = losscode;
                        histLossDetails.MessageCodeId = losscodeid;
                        histLossDetails.MessageDesc = losscode;
                        db.Entry(histLossDetails).State = EntityState.Modified;
                        db.SaveChanges();
                    }
                    //row.IsUpdate = 1;
                    //db.Entry(row).State = EntityState.Modified;
                    //db.SaveChanges();

                    //DateTime stTime = Convert.ToDateTime(tcfdata.CorrectedDate + " " + "00:00:00");
                    //DateTime etTime = Convert.ToDateTime(tcfdata.CorrectedDate + " " + "00:00:00");                  

                }

                DateTime stTime = Convert.ToDateTime(correctedDate);
                DateTime etTime = Convert.ToDateTime(correctedDate);

                check = TakeBackupReportData(correctedDate);

                log.Error("Before Method");
                log.Error("Before Method oee:" + check);
                //bool calres = CalculateOEEForYesterday(stTime, etTime);
                if (check)
                {
                    log.Error("after Method" +check);
                    DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
                    List<int> oeecal = new List<int>();
                    oeecal.Add(MachineID);
                    Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(stTime, etTime,oeecal);  // for WO report updation
                    log.Error("After reportWOUpdate:" + reportWOUpdate.Result);
                    Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(stTime, etTime,oeecal);// for OEE report updation
                    log.Error("After reportOEEUpdate:" + reportOEEUpdate.Result);
                    if (reportOEEUpdate.Result == true && reportOEEUpdate.Result == true)
                    {
                        log.Error("   if (reportOEEUpdate.Result == true && reportOEEUpdate.Result == true)");
                        log.Error("After if condition:" + reportOEEUpdate.Result);
                        foreach (var row in tcfrecord)
                        {
                            row.IsUpdate = 1;
                            db.SaveChanges();
                            entity.isTrue = true;
                            entity.response = ResourceResponse.SuccessMessage;
                        }
                    }
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.FailureMessage;
                }


            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error("exception");
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }


        public void CheckSplitDurationLossUpdation(string correctedDate)
        {
            try
            {

                var lossOfEnteyRecord = db.Tbltcflossofentry.Where(m => (m.IsAccept == 1 || m.IsAccept1 == 1) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate).OrderBy(x => x.LossId).Select(x => x.LossId).ToList().Distinct();
                if (lossOfEnteyRecord.Count() > 0)
                {
                    foreach (int lossId in lossOfEnteyRecord)
                    {
                        var lossMultipleRecord = db.Tbltcflossofentry.Where(m => (m.IsAccept == 1 || m.IsAccept1 == 1) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate && m.LossId == lossId).ToList();
                        if (lossMultipleRecord.Count() > 1)
                        {
                            int i = 0;
                            //lossMultipleRecord = lossMultipleRecord.Skip(1).ToList();
                            foreach (var rowLoss in lossMultipleRecord)
                            {
                                string losscode = db.Tbllossescodes.Where(m => m.LossCodeId == rowLoss.MessageCodeId).Select(m => m.LossCode).FirstOrDefault();
                                if (i == 0)
                                {
                                    var livelossdet = db.Tbllivelossofentry.Where(m => m.LossId == lossId).FirstOrDefault();
                                    var histLossDetails = db.Tbllossofentry.Where(m => m.LossId == lossId).FirstOrDefault();
                                    if (livelossdet != null)
                                    {
                                        livelossdet.MessageCode = losscode;
                                        livelossdet.MessageCodeId = Convert.ToInt32(rowLoss.MessageCodeId);
                                        livelossdet.MessageDesc = losscode;
                                        livelossdet.EndDateTime = rowLoss.EndDateTime;
                                        db.Entry(livelossdet).State = EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                    if (histLossDetails != null)
                                    {
                                        histLossDetails.MessageCode = losscode;
                                        histLossDetails.MessageCodeId = Convert.ToInt32(rowLoss.MessageCodeId);
                                        histLossDetails.MessageDesc = losscode;
                                        histLossDetails.EndDateTime = rowLoss.EndDateTime;
                                        db.Entry(histLossDetails).State = EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                    i++;
                                }
                                else
                                {
                                    string shift = GetShift(Convert.ToDateTime(rowLoss.EndDateTime));

                                    Tbllivelossofentry addRowLive = new Tbllivelossofentry();
                                    addRowLive.CorrectedDate = rowLoss.CorrectedDate;
                                    addRowLive.DoneWithRow = 1;
                                    addRowLive.EndDateTime = rowLoss.EndDateTime;
                                    addRowLive.EntryTime = rowLoss.StartDateTime;
                                    addRowLive.ForRefresh = 0;
                                    addRowLive.IsScreen = 0;
                                    addRowLive.IsStart = 0;
                                    addRowLive.IsUpdate = 0;
                                    addRowLive.MachineId = Convert.ToInt32(rowLoss.MachineId);
                                    addRowLive.MessageCode = losscode;
                                    addRowLive.MessageCodeId = Convert.ToInt32(rowLoss.MessageCodeId);
                                    addRowLive.MessageDesc = losscode;
                                    addRowLive.Shift = shift;
                                    addRowLive.StartDateTime = rowLoss.StartDateTime;
                                    db.Tbllivelossofentry.Add(addRowLive);
                                    db.SaveChanges();
                                    int liveLossId = addRowLive.LossId;

                                    Tbllossofentry addRow = new Tbllossofentry();
                                    addRow.CorrectedDate = rowLoss.CorrectedDate;
                                    addRow.DoneWithRow = 1;
                                    addRow.EndDateTime = rowLoss.EndDateTime;
                                    addRow.EntryTime = rowLoss.StartDateTime;
                                    addRow.ForRefresh = 0;
                                    addRow.IsScreen = 0;
                                    addRow.IsStart = 0;
                                    addRow.IsUpdate = 0;
                                    addRow.MachineId = Convert.ToInt32(rowLoss.MachineId);
                                    addRow.MessageCode = losscode;
                                    addRow.MessageCodeId = Convert.ToInt32(rowLoss.MessageCodeId);
                                    addRow.MessageDesc = losscode;
                                    addRow.Shift = shift;
                                    addRow.StartDateTime = rowLoss.StartDateTime;
                                    addRow.LossId = liveLossId;
                                    db.Tbllossofentry.Add(addRow);
                                    db.SaveChanges();

                                    rowLoss.LossId = liveLossId;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
        }

        public string GetShift(DateTime DateNow)
        {
            string ShiftValue = "C";
            try
            {
                //DateTime DateNow = DateTime.Now;
                var ShiftDetails = db.TblshiftMstr.Where(m => m.IsDeleted == 0).ToList();
                foreach (var row in ShiftDetails)
                {
                    int ShiftStartHour = row.StartTime.Value.Hours;
                    int ShiftEndHour = row.EndTime.Value.Hours;
                    int CurrentHour = DateNow.Hour;
                    if (CurrentHour >= ShiftStartHour && CurrentHour <= ShiftEndHour)
                    {
                        ShiftValue = row.ShiftName;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }

            return ShiftValue;
        }


        public bool CheckSplitDurationNeedUpdation()
        {
            bool result = false;
            try
            {

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }


        //when they enter multiple reasons insertion into tcf table
        //public EntityModel InsertMultiLoss(tcfclass tcfdata, int noofreason)
        //{
        //    EntityModel entity = new EntityModel();
        //    try
        //    {
        //        for (int i = 0; i < noofreason; i++)
        //        {
        //            var tcfrecord = db.Tbltcflossofentry.Where(m => m.Ncid == tcfdata.NCID).FirstOrDefault();
        //            try
        //            {
        //                if (noofreason == 1)
        //                {
        //                    if (tcfrecord != null)
        //                    {

        //                        //updating the losscodes for previous loss
        //                        tcfrecord.EndDateTime = tcfdata.EndTime;
        //                        tcfrecord.StartDateTime = tcfrecord.StartDateTime;
        //                        tcfrecord.ReasonLevel1 = Convert.ToString(tcfdata.reasonLevel1Id);
        //                        tcfrecord.ReasonLevel2 = Convert.ToString(tcfdata.reasonLevel2Id);
        //                        tcfrecord.ReasonLevel3 = Convert.ToString(tcfdata.reasonLevel3Id);
        //                        db.Entry(tcfrecord).State = EntityState.Modified;
        //                        db.SaveChanges();
        //                    }
        //                }
        //                else if (noofreason > 1)
        //                {
        //                    //Inserting one new record for that loss
        //                    Tbltcflossofentry tcfobj = new Tbltcflossofentry();
        //                    tcfobj.LossId = tcfrecord.LossId;
        //                    tcfobj.MachineId = tcfrecord.MachineId;
        //                    tcfobj.MessageCodeId = tcfrecord.MessageCodeId;
        //                    tcfrecord.ReasonLevel1 = Convert.ToString(tcfdata.reasonLevel1Id);
        //                    tcfrecord.ReasonLevel2 = Convert.ToString(tcfdata.reasonLevel2Id);
        //                    tcfrecord.ReasonLevel3 = Convert.ToString(tcfdata.reasonLevel3Id);
        //                    tcfobj.StartDateTime = tcfdata.StartTime;
        //                    tcfobj.EndDateTime = tcfrecord.EndDateTime;
        //                    tcfobj.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
        //                    tcfobj.IsUpdate = 0;
        //                    tcfobj.IsAccept = 0;
        //                    db.Tbltcflossofentry.Add(tcfobj);
        //                    db.SaveChanges();
        //                    entity.isTrue = true;
        //                    entity.response = ResourceResponse.SuccessMessage;
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                entity.isTrue = false;
        //                entity.response = ResourceResponse.ExceptionMessage;
        //                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //            }
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return entity;

        //}

        //when click on set to approval button  

        public EntityModel SendToApproveAllLossDetails(EntityHMIDetails data)
        {
            EntityModel obj = new EntityModel();   // get all the details then send by url or by html
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
                    machName = "No Code " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
                }

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\ReasonEmailTemplate1.html");
                string htmlStr = File.ReadAllText(reader);

                string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";
                String[] seperator = { "{{reasonStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var reasonHtml = htmlArr[1].Split(new String[] { "{{reasonEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{reasonStart}}", "");
                htmlStr = htmlStr.Replace("{{reasonEnd}}", "");
                string CorrectedDate = "";
                int sl = 1;

                foreach (var machineRow in machineData)
                {
                    int i = 0;
                    int machineId = machineRow.MachineId;

                    var tcfloss = db.Tbltcflossofentry.Where(x => x.IsArroval == 0 && x.IsAccept == 0 && x.IsUpdate == 0 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.ReasonLevel1 != null).OrderBy(m=>m.Ncid).ToList();
                    if (tcfloss.Count > 0)
                    {

                        foreach (var row in tcfloss)
                        {
                            row.IsArroval = 1;
                            row.IsHold = 1;
                            db.SaveChanges();

                            string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => x.MachineInvNo).FirstOrDefault();

                            CorrectedDate = row.CorrectedDate;
                            String slno = Convert.ToString(sl);
                            String Stime = Convert.ToString(row.StartDateTime);
                            String Etime = Convert.ToString(row.EndDateTime);
                            int losscodeid1 = Convert.ToInt32(row.ReasonLevel1);
                            String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                            int losscodeid2 = Convert.ToInt32(row.ReasonLevel2);
                            String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                            int losscodeid3 = Convert.ToInt32(row.ReasonLevel3);
                            String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{machineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{rl1}}", losscode1);
                            htmlStr = htmlStr.Replace("{{rl2}}", losscode2);
                            htmlStr = htmlStr.Replace("{{rl3}}", losscode3);
                            htmlStr = htmlStr.Replace("{{startTime}}", Stime);
                            htmlStr = htmlStr.Replace("{{endTime}}", Etime);
                            if (tcfloss.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", "");
                            }
                            else if (sl < tcfloss.Count)
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                            }
                            sl++;
                        }
                    }
                }
                htmlStr = htmlStr.Replace(reasonHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "For 1st Level Approval");

                // string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLNoCode").Value;
                string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLNoCode").Value;

                //String rejectSrc = @"http://192.168.0.15:8082/reasonacceptreject?CorrectDate=" + CorrectedDate + "&id=0";
                //String acceptSrc = @"http://192.168.0.15:8082/reasonacceptreject?CorrectDate=" + CorrectedDate + "&id=1";

                string rejectSrc = rejectUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "&checked=0";

                //string acceptSrc = acceptUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";


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


                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.ShopId == shopId).ToList();
                }
                foreach (var row in tcfApproveMail)
                {
                    toMailIds += row.FirstApproverToList + ",";
                    ccMailIds += row.FirstApproverCcList + ",";
                }

                htmlStr = htmlStr.Replace("{{reason}}", "");
                htmlStr = htmlStr.Replace("{{userName}}", toName);
                //htmlStr = htmlStr.Replace("{{Sname}}", "Saurabh");
                //htmlStr = htmlStr.Replace("{{Lurl}}", logo);
                //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                toMailIds = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);// removing last comma

                //bool check = SendMail(htmlStr, toMailID, ccMailID, 1);

                bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 1, machName);

                if (ret)
                {
                    obj.isTrue = true;
                    obj.response = "Send For Approval";
                }
                else
                {
                    obj.isTrue = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }

            catch (Exception ex)
            {
                obj.isTrue = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //public EntityModel GetRejecttcfLoss(int id)
        //{
        //    EntityModel entity = new EntityModel();
        //    try
        //    {
        //        var tcfrecord = db.Tbltcflossofentry.Where(m => m.IsArroval == 1).ToList();
        //        foreach (var row in tcfrecord)
        //        {
        //            row.IsAccept = 2;
        //            row.RejectReasonId = id;
        //            db.Entry(row).State = EntityState.Modified;
        //            db.SaveChanges();
        //            entity.isTrue = true;
        //            entity.response = "Reject Reason Updated Successfully";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        entity.isTrue = false;
        //        entity.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return entity;
        //}

        public EntityModel GetReason()
        {
            EntityModel entity = new EntityModel();

            try
            {
                List<RejectReason> Reject = new List<RejectReason>();
                var reason = db.Tblrejectreason.Where(m => m.IsDeleted == 0 && m.IsTcf == 1).ToList();
                foreach (var row in reason)
                {
                    RejectReason objrea = new RejectReason();
                    objrea.RejectID = row.Rid;
                    objrea.RejectName = row.RejectName;
                    Reject.Add(objrea);
                }
                if (Reject != null)
                {
                    entity.isTrue = true;
                    entity.response = Reject;
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        public CommonResponse AcceptAllNoCodeDetails(EntityHMIDetails data)
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
                    machName = "No Code " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
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
                        int ncid = Convert.ToInt32(idrow);
                        var getNoCodeDet = db.Tbltcflossofentry.Where(m => m.MachineId == machine && (m.IsAccept == 0 || m.IsAccept1 == 0) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate && m.Ncid == ncid).OrderBy(m=>m.Ncid).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            //foreach (var row in getNoCodeDet)
                            //{
                            //row.IsAccept = 1;
                            //db.SaveChanges();

                            if (getNoCodeDet.IsAccept == 0)
                            {
                                getNoCodeDet.IsAccept = 1;
                                getNoCodeDet.IsHold = 0;
                                getNoCodeDet.ApprovalLevel = 1;
                                db.SaveChanges();
                                if (getNoCodeDet.UpdateLevel == 0)
                                {
                                    getNoCodeDet.UpdateLevel = 1;
                                    db.SaveChanges();
                                }
                                if (getNoCodeDet.UpdateLevel == 1)
                                {
                                    updateReport = true;
                                    appLevel = 1;
                                }
                                else
                                {
                                    appLevel = 2;
                                }
                            }
                            else if (getNoCodeDet.IsAccept1 == 0)
                            {
                                getNoCodeDet.IsAccept1 = 1;
                                getNoCodeDet.IsHold = 0;
                                getNoCodeDet.ApprovalLevel = 2;
                                db.SaveChanges();
                                if (getNoCodeDet.UpdateLevel == 0)
                                {
                                    getNoCodeDet.UpdateLevel = 2;
                                    db.SaveChanges();
                                }
                                if (getNoCodeDet.UpdateLevel == 2)
                                {
                                    updateReport = true;
                                }
                                //appLevel = 2;
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
                            var getNoCodeDet = db.Tbltcflossofentry.Where(m => m.Ncid == id).FirstOrDefault();
                            if (getNoCodeDet != null)
                            {
                                getNoCodeDet.IsHold = 1;
                                db.SaveChanges();
                            }
                        }
                    }

                }

                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.ShopId == shopId).ToList();
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

                toMailIds = toMailIds.Remove(toMailIds.Length - 1);
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);


                if (updateReport)
                {
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\ReasonEmailTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);

                    string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";
                    String[] seperator = { "{{reasonStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var reasonHtml = htmlArr[1].Split(new String[] { "{{reasonEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{reasonStart}}", "");
                    htmlStr = htmlStr.Replace("{{reasonEnd}}", "");
                    string CorrectedDate = "";
                    int sl = 1;

                    foreach (var machineRow in machineData)
                    {
                        int i = 0;
                        int machineId = machineRow.MachineId;

                        var tcfloss = db.Tbltcflossofentry.Where(x => x.IsArroval == 1 && x.IsAccept == 1 && x.IsAccept1 == 1 && x.IsUpdate == 0 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.ReasonLevel1 != null).OrderBy(m => m.Ncid).ToList();
                        if (tcfloss.Count > 0)
                        {
                            // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");

                            foreach (var row in tcfloss)
                            {
                                //row.IsArroval = 1;
                                //db.SaveChanges();

                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => x.MachineInvNo).FirstOrDefault();

                                CorrectedDate = row.CorrectedDate;
                                String slno = Convert.ToString(sl);
                                String Stime = Convert.ToString(row.StartDateTime);
                                String Etime = Convert.ToString(row.EndDateTime);
                                int losscodeid1 = Convert.ToInt32(row.ReasonLevel1);
                                String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid2 = Convert.ToInt32(row.ReasonLevel2);
                                String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid3 = Convert.ToInt32(row.ReasonLevel3);
                                String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                                htmlStr = htmlStr.Replace("{{slno}}", slno);
                                htmlStr = htmlStr.Replace("{{machineName}}", machineName);
                                htmlStr = htmlStr.Replace("{{rl1}}", losscode1);
                                htmlStr = htmlStr.Replace("{{rl2}}", losscode2);
                                htmlStr = htmlStr.Replace("{{rl3}}", losscode3);
                                htmlStr = htmlStr.Replace("{{startTime}}", Stime);
                                htmlStr = htmlStr.Replace("{{endTime}}", Etime);
                                if (tcfloss.Count == 1)
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", "");
                                }
                                else if (sl < tcfloss.Count)
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                                }
                                else
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                                }
                                sl++;
                            }
                        }
                    }
                    htmlStr = htmlStr.Replace(reasonHtml, "");
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "These Reasons are Accepted");

                    //string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLNoCode").Value;

                    //string rejectSrc = rejectUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";


                    //htmlStr = htmlStr.Replace("{{reason}}", "");
                    //htmlStr = htmlStr.Replace("{{userName}}", toName);
                    //htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                    bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 2, machName);

                    if (ret)
                    {
                        obj.isTure = true;
                        obj.response = "Mail Sent";
                        UpdateliveLoss(correctedDate,data.machineId);
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = ResourceResponse.FailureMessage;
                    }

                    //string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                    //"</head><body><p>Dear All,</p></br><p><center> The LossCode Reasons Has Been Accepted</center></p></br><p>Thank you" +
                    //"</p></body></html>";

                    //bool ret = SendMail(message, toMailIds, ccMailIds, 0, machName);
                    //if (ret)
                    //{
                    //    obj.isTure = true;
                    //    obj.response = ResourceResponse.SuccessMessage;
                    //    UpdateliveLoss(correctedDate);
                    //}
                    //else
                    //{
                    //    obj.isTure = false;
                    //    obj.response = ResourceResponse.FailureMessage;
                    //}
                }
                else
                {
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\ReasonEmailTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);

                    string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";
                    String[] seperator = { "{{reasonStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var reasonHtml = htmlArr[1].Split(new String[] { "{{reasonEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{reasonStart}}", "");
                    htmlStr = htmlStr.Replace("{{reasonEnd}}", "");
                    string CorrectedDate = "";
                    int sl = 1;

                    foreach (var machineRow in machineData)
                    {
                        int i = 0;
                        int machineId = machineRow.MachineId;

                        var tcfloss = db.Tbltcflossofentry.Where(x => x.IsArroval == 1 && x.IsAccept == 1 && x.IsAccept1 == 0 && x.IsUpdate == 0 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.ReasonLevel1 != null).OrderBy(m => m.Ncid).ToList();
                        if (tcfloss.Count > 0)
                        {
                            // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");

                            foreach (var row in tcfloss)
                            {
                                //row.IsArroval = 1;
                                //db.SaveChanges();

                                string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => x.MachineInvNo).FirstOrDefault();

                                CorrectedDate = row.CorrectedDate;
                                String slno = Convert.ToString(sl);
                                String Stime = Convert.ToString(row.StartDateTime);
                                String Etime = Convert.ToString(row.EndDateTime);
                                int losscodeid1 = Convert.ToInt32(row.ReasonLevel1);
                                String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid2 = Convert.ToInt32(row.ReasonLevel2);
                                String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid3 = Convert.ToInt32(row.ReasonLevel3);
                                String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                                htmlStr = htmlStr.Replace("{{slno}}", slno);
                                htmlStr = htmlStr.Replace("{{machineName}}", machineName);
                                htmlStr = htmlStr.Replace("{{rl1}}", losscode1);
                                htmlStr = htmlStr.Replace("{{rl2}}", losscode2);
                                htmlStr = htmlStr.Replace("{{rl3}}", losscode3);
                                htmlStr = htmlStr.Replace("{{startTime}}", Stime);
                                htmlStr = htmlStr.Replace("{{endTime}}", Etime);
                                if (tcfloss.Count == 1)
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", "");
                                }
                                else if (sl < tcfloss.Count)
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                                }
                                else
                                {
                                    htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                                }
                                sl++;
                            }
                        }
                    }
                    htmlStr = htmlStr.Replace(reasonHtml, "");
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "For 2nd Level Approval");

                    //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLNoCode").Value;
                    string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLNoCode").Value;

                    //string rejectSrc = rejectUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";

                    string rejectSrc = rejectUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "&checked=" + data.id + "";
                    //string acceptSrc = acceptUrl + "correctedDate=" + CorrectedDate + "&plantId=" + plantId + "&shopId=" + shopId + "&cellId=" + cellId + "&machineId=" + machId + "";


                    htmlStr = htmlStr.Replace("{{reason}}", "");
                    htmlStr = htmlStr.Replace("{{userName}}", toName);
                    //htmlStr = htmlStr.Replace("{{Sname}}", "Saurabh");
                    //htmlStr = htmlStr.Replace("{{Lurl}}", logo);
                    //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                    htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                    bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 1, machName);

                    if (ret)
                    {
                        obj.isTure = true;
                        obj.response = "Send For 2nd Approval";
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

            }
            return obj;
        }

        public CommonResponse RejectAllNoCodeDetails(RejectReasonStore data)
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
                    machName = "No Code " + db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machId).Select(x => x.MachineInvNo).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.cellId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => x.CellName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.shopiId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).Select(x => x.ShopName).FirstOrDefault() + " " + correctedDate;
                }
                else if (data.plantId != 0)
                {
                    machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                    machName = "No Code " + db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).Select(x => x.PlantName).FirstOrDefault() + " " + correctedDate;
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
                        int ncid = Convert.ToInt32(idrow);
                        var getNoCodeDet = db.Tbltcflossofentry.Where(m => m.MachineId == machine && (m.IsAccept == 0 || m.IsAccept1 == 0) && m.IsUpdate == 0 && m.IsArroval == 1 && m.CorrectedDate == correctedDate && m.ReasonLevel1 != null && m.Ncid == ncid).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            //foreach (var row in getNoCodeDet)
                            //{
                            //row.IsAccept = 2;
                            //row.RejectReasonId = reasonId;
                            //db.SaveChanges();

                            if (getNoCodeDet.IsAccept == 0)
                            {
                                getNoCodeDet.IsAccept = 2;
                                getNoCodeDet.IsHold = 0;
                                getNoCodeDet.RejectReasonId = reasonId;
                                getNoCodeDet.ApprovalLevel = 1;
                                db.SaveChanges();
                                checkMail = 1;
                            }
                            else if (getNoCodeDet.IsAccept1 == 0)
                            {
                                getNoCodeDet.IsAccept1 = 2;
                                getNoCodeDet.IsHold = 0;
                                getNoCodeDet.RejectReasonId1 = reasonId;
                                getNoCodeDet.ApprovalLevel = 2;
                                db.SaveChanges();
                                checkMail = 2;
                            }

                        }

                        //}
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
                            var getNoCodeDet = db.Tbltcflossofentry.Where(m => m.Ncid == id).FirstOrDefault();
                            if (getNoCodeDet != null)
                            {
                                getNoCodeDet.IsHold = 1;
                                db.SaveChanges();
                            }
                        }
                    }
                }

                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.CellId == cellId).ToList();
                if (tcfApproveMail.Count() == 0)
                {
                    tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 1 && x.ShopId == shopId).ToList();
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

                toMailIds = toMailIds.Remove(toMailIds.Length - 1);
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);

                string reasonName = db.Tblrejectreason.Where(x => x.IsDeleted == 0 && x.Rid == reasonId).Select(x => x.RejectNameDesc).FirstOrDefault();

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\ReasonEmailTemplate1.html");
                string htmlStr = File.ReadAllText(reader);

                string logo = @"C:\TataReport\TCFTemplate\120px-Tata_logo.Jpeg";
                String[] seperator = { "{{reasonStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var reasonHtml = htmlArr[1].Split(new String[] { "{{reasonEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{reasonStart}}", "");
                htmlStr = htmlStr.Replace("{{reasonEnd}}", "");
                string CorrectedDate = "";
                int sl = 1;

                foreach (var machineRow in machineData)
                {
                    int i = 0;
                    int machineId = machineRow.MachineId;

                    var tcfloss = db.Tbltcflossofentry.Where(x => x.IsArroval == 1 && x.IsAccept == 2 && x.IsAccept1 == 2 && x.IsUpdate == 0 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.ReasonLevel1 != null).OrderBy(m => m.Ncid).ToList();
                    if (tcfloss.Count > 0)
                    {
                        // var reader = Path.Combine(@"D:\Monika\TCF\TCF\ReasonEmailTemplate.html");

                        foreach (var row in tcfloss)
                        {
                            //row.IsArroval = 1;
                            //db.SaveChanges();

                            string machineName = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => x.MachineInvNo).FirstOrDefault();

                            CorrectedDate = row.CorrectedDate;
                            String slno = Convert.ToString(sl);
                            String Stime = Convert.ToString(row.StartDateTime);
                            String Etime = Convert.ToString(row.EndDateTime);
                            int losscodeid1 = Convert.ToInt32(row.ReasonLevel1);
                            String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                            int losscodeid2 = Convert.ToInt32(row.ReasonLevel2);
                            String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                            int losscodeid3 = Convert.ToInt32(row.ReasonLevel3);
                            String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{machineName}}", machineName);
                            htmlStr = htmlStr.Replace("{{rl1}}", losscode1);
                            htmlStr = htmlStr.Replace("{{rl2}}", losscode2);
                            htmlStr = htmlStr.Replace("{{rl3}}", losscode3);
                            htmlStr = htmlStr.Replace("{{startTime}}", Stime);
                            htmlStr = htmlStr.Replace("{{endTime}}", Etime);
                            if (tcfloss.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", "");
                            }
                            else if (sl < tcfloss.Count)
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                            }
                            else
                            {
                                htmlStr = htmlStr.Replace("{{reason}}", reasonHtml);
                            }
                            sl++;
                        }
                    }
                }
                htmlStr = htmlStr.Replace(reasonHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "These Loss Reasons Has Been Rejected for this " + reasonName + ".");


               // string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
               //"</head><body><p>Dear " + toName + ",</p></br><p><center> The Loss Reasons Has Been Rejected for this " + reasonName + ". </center></p></br><p>Thank you" +
               //"</p></body></html>";

                bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 2, machName);

                if (ret)
                {
                    obj.isTure = true;
                    obj.response = ResourceResponse.SuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = ResourceResponse.FailureMessage;
                }
            }
            catch (Exception ex)
            {

            }
            return obj;
        }

        public CommonResponse AcceptRejectNoCodeTable(EntityHMIDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<getReasonTable> reaobjlist = new List<getReasonTable>();
                string correctedDate = data.fromDate;
                int cellId = data.cellId;
                int machId = data.machineId;
                int shopId = data.shopiId;
                int plantId = data.plantId;
                string[] ids = data.id.Split(',');
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
                        //var NoCodeDet = db.Tbltcflossofentry.Where(x => x.IsUpdate == 0 && x.IsArroval == 1 && x.CorrectedDate == correctedDate && (x.IsAccept == 0 || x.IsAccept == 1) && (x.IsAccept1 == 1 || x.IsAccept == 0) && x.MachineId == machine && x.ReasonLevel1 != null).ToList();
                        var NoCodeDet = db.Tbltcflossofentry.Where(x => x.IsUpdate == 0 && x.IsArroval == 1 && x.CorrectedDate == correctedDate && x.IsHold == 1 && x.MachineId == machine && x.ReasonLevel1 != null).OrderBy(m => m.Ncid).ToList();
                        if (NoCodeDet.Count > 0)
                        {
                            foreach (var row in NoCodeDet)
                            {
                                var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == row.MachineId).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                getReasonTable reaobj = new getReasonTable();
                                String Stime = Convert.ToString(row.StartDateTime);
                                String Etime = Convert.ToString(row.EndDateTime);
                                int losscodeid1 = Convert.ToInt32(row.ReasonLevel1);
                                String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid2 = Convert.ToInt32(row.ReasonLevel2);
                                String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                                int losscodeid3 = Convert.ToInt32(row.ReasonLevel3);
                                String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                                reaobj.ncid = row.Ncid;
                                reaobj.MachineName = machineRow.MachineInvNo;
                                reaobj.reasonLevel1 = losscode1;
                                reaobj.reasonLevel2 = losscode2;
                                reaobj.reasonLevel3 = losscode3;
                                reaobj.StartTime = Convert.ToString(Stime);
                                reaobj.EndTime = Convert.ToString(Etime);
                                reaobjlist.Add(reaobj);
                            }
                            obj.isTure = true;
                            obj.response = reaobjlist.OrderBy(x => x.StartTime);
                        }
                        else if (NoCodeDet.Count == 0)
                        {
                            obj.isTure = false;
                            obj.response = "All The Reasons are Accepted";
                        }
                    }
                    else
                    {
                        bool ret = false;
                        foreach (var rowid in ids)
                        {
                            int id = Convert.ToInt32(rowid);
                            var nocdet = db.Tbltcflossofentry.Where(x => x.Ncid == id && x.IsHold == 0).FirstOrDefault();
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
                                var NoCodeDet = db.Tbltcflossofentry.Where(x => x.Ncid == id && x.IsHold == 0 && (x.IsAccept == 1 || x.IsAccept1 == 1) && x.ApprovalLevel !=2).FirstOrDefault();
                                if (NoCodeDet != null)
                                {

                                    //foreach (var row in NoCodeDet)
                                    //{
                                    var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == NoCodeDet.MachineId).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                    getReasonTable reaobj = new getReasonTable();
                                    String Stime = Convert.ToString(NoCodeDet.StartDateTime);
                                    String Etime = Convert.ToString(NoCodeDet.EndDateTime);
                                    int losscodeid1 = Convert.ToInt32(NoCodeDet.ReasonLevel1);
                                    String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                                    int losscodeid2 = Convert.ToInt32(NoCodeDet.ReasonLevel2);
                                    String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                                    int losscodeid3 = Convert.ToInt32(NoCodeDet.ReasonLevel3);
                                    String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                                    reaobj.ncid = NoCodeDet.Ncid;
                                    reaobj.MachineName = machineRow.MachineInvNo;
                                    reaobj.reasonLevel1 = losscode1;
                                    reaobj.reasonLevel2 = losscode2;
                                    reaobj.reasonLevel3 = losscode3;
                                    reaobj.StartTime = Convert.ToString(Stime);
                                    reaobj.EndTime = Convert.ToString(Etime);
                                    reaobjlist.Add(reaobj);
                                    //}
                                    obj.isTure = true;
                                    obj.response = reaobjlist.OrderBy(x => x.StartTime);
                                }

                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "All The Reasons are Accepted";
                                }
                            }
                        }
                        else
                        {
                            foreach (var rowid in ids)
                            {
                                int id = Convert.ToInt32(rowid);
                                var NoCodeDet = db.Tbltcflossofentry.Where(x => x.Ncid == id && x.IsHold == 1).FirstOrDefault();
                                if (NoCodeDet != null)
                                {
                                    //foreach (var row in NoCodeDet)
                                    //{
                                    var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == NoCodeDet.MachineId).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                                    getReasonTable reaobj = new getReasonTable();
                                    String Stime = Convert.ToString(NoCodeDet.StartDateTime);
                                    String Etime = Convert.ToString(NoCodeDet.EndDateTime);
                                    int losscodeid1 = Convert.ToInt32(NoCodeDet.ReasonLevel1);
                                    String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                                    int losscodeid2 = Convert.ToInt32(NoCodeDet.ReasonLevel2);
                                    String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                                    int losscodeid3 = Convert.ToInt32(NoCodeDet.ReasonLevel3);
                                    String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                                    reaobj.ncid = NoCodeDet.Ncid;
                                    reaobj.MachineName = machineRow.MachineInvNo;
                                    reaobj.reasonLevel1 = losscode1;
                                    reaobj.reasonLevel2 = losscode2;
                                    reaobj.reasonLevel3 = losscode3;
                                    reaobj.StartTime = Convert.ToString(Stime);
                                    reaobj.EndTime = Convert.ToString(Etime);
                                    reaobjlist.Add(reaobj);
                                    //}
                                    obj.isTure = true;
                                    obj.response = reaobjlist.OrderBy(x => x.StartTime);
                                }
                            }
                        }


                        //var NoCodeDet1 = db.Tbltcflossofentry.Where(x => x.IsHold == 1 && x.IsAccept1 == 0 && x.ApprovalLevel==1).ToList();
                        //    if (NoCodeDet1.Count > 0)
                        //    {
                        //        foreach (var norow in NoCodeDet1)
                        //        {
                        //            var machineDetais = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == norow.MachineId).Select(x => new { x.MachineInvNo, x.IsNormalWc }).FirstOrDefault();
                        //            getReasonTable reaobj = new getReasonTable();
                        //            String Stime = Convert.ToString(norow.StartDateTime);
                        //            String Etime = Convert.ToString(norow.EndDateTime);
                        //            int losscodeid1 = Convert.ToInt32(norow.ReasonLevel1);
                        //            String losscode1 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid1).Select(m => m.LossCode).FirstOrDefault();
                        //            int losscodeid2 = Convert.ToInt32(norow.ReasonLevel2);
                        //            String losscode2 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid2).Select(m => m.LossCode).FirstOrDefault();
                        //            int losscodeid3 = Convert.ToInt32(norow.ReasonLevel3);
                        //            String losscode3 = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid3).Select(m => m.LossCode).FirstOrDefault();
                        //            reaobj.ncid = norow.Ncid;
                        //            reaobj.MachineName = machineRow.MachineInvNo;
                        //            reaobj.reasonLevel1 = losscode1;
                        //            reaobj.reasonLevel2 = losscode2;
                        //            reaobj.reasonLevel3 = losscode3;
                        //            reaobj.StartTime = Convert.ToString(Stime);
                        //            reaobj.EndTime = Convert.ToString(Etime);
                        //            reaobjlist.Add(reaobj);
                        //            //}
                        //            obj.isTure = true;
                        //            obj.response = reaobjlist.OrderBy(x => x.StartTime);
                        //        }

                        //    }
                        //    else
                        //    {
                        //        obj.isTure = false;
                        //        obj.response = "All The Reasons are Accepted";
                        //    }
                        //}


                    }
                }
            }
            catch (Exception ex)
            {

            }
            return obj;
        }

        #region MachineDetails

        //public EntityModel GetMachineDet(int cellid)
        //{
        //    EntityModel entiy = new EntityModel();

        //    List<Machine> machinelist = new List<Machine>();
        //    var machinedet = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellId == cellid).ToList();
        //    foreach (var item in machinedet)
        //    {
        //        Machine macobj = new Machine();
        //        macobj.Machineid = item.MachineId;
        //        macobj.MachinedispNmae = item.MachineDispName;
        //        machinelist.Add(macobj);
        //    }
        //    if (machinelist != null)
        //    {
        //        entiy.isTrue = true;
        //        entiy.response = machinelist;
        //    }
        //    else
        //    {
        //        entiy.isTrue = false;
        //        entiy.response = "fail";
        //    }
        //    return entiy;
        //}
        #endregion

        public bool SendMail(string message, string tolist, string cclist, int image, string subject)
        {
            bool ret = false;

            log.Error(tolist);

            try
            {
                if (message != "" && tolist != "")
                {
                    MailMessage mail = new MailMessage();
                    mail.To.Add(tolist);
                    if (cclist != "")
                    {
                        mail.CC.Add(cclist);
                    }

                    var smtpConn = db.Smtpdetails.Where(x => x.IsDeleted == true && x.TcfModuleId == 1).FirstOrDefault();
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
                    //mail.From = new MailAddress("unitworks@tasl.aero");
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
                    else if(image == 2)
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
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return ret;
        }

        //Report Backup Data
        public bool TakeBackupReportData(string correctedDate)
        {
            bool result = false;
            try
            {
                //getting the connection string from app string.json
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                string dbName = configuration.GetSection("MySettings").GetSection("Schema").Value;
                DataTable dt = new DataTable();
                string queryUnAssignedMachine = "select Distinct(Machineid) from " + dbName + ".[tbltcflossofentry] where Correcteddate='" + correctedDate + "' and (IsAccept=1 or IsAccept1=1) and Isupdate=0";
                SqlDataAdapter sDA = new SqlDataAdapter(queryUnAssignedMachine, connectionString);
                sDA.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        bool oeeCheck = false;
                        DateTime corrDate = Convert.ToDateTime(correctedDate + " " + "00:00:00");
                        int machineId = Convert.ToInt32(dt.Rows[i][0]);
                        var oeeDasjboardVar = db.Tbloeedashboardvariables.Where(x => x.IsDeleted == 0 && x.Wcid == machineId && x.StartDate == corrDate).FirstOrDefault();
                        if(oeeDasjboardVar == null)
                        {
                            oeeCheck = true;
                        }
                        else
                        {
                            oeeCheck = InsertToOEEDashboardVar(oeeDasjboardVar);
                        }
                        var woreport = db.Tblworeport.Where(x => x.CorrectedDate == correctedDate && x.MachineId == machineId).ToList();
                        bool woReportCheck = InsertToWoReport(woreport);
                        var woLoss = db.Tblhmiscreen.Where(x => x.MachineId == machineId && x.CorrectedDate == correctedDate).Select(x => x.Hmiid).ToList();
                        bool woLossCheck = InsertToWoLoss(woLoss);
                        log.Error("OEE" + oeeCheck + "  WOReport" + woReportCheck + " WOLoss" + woLossCheck);
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
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
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
                    TblBackupoeedashboardvariables addRow = new TblBackupoeedashboardvariables();
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
                    db.TblBackupoeedashboardvariables.Add(addRow);
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
                            result = false;
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

        /// <summary>
        /// No Code Split Duration Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 NoCodeSplitDurationDetails(List<NoCodeStartEndDateTime> data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                int i = 0;
                List<NoCodeStartEndDateTime> modeStartEndDateTimeList = new List<NoCodeStartEndDateTime>();
                foreach (var item in data)
                {
                    i++;
                    modeStartEndDateTimeList.Add(item);
                    if (data.Count == i)
                    {
                        var dbCheck = db.Tbltcflossofentry.Where(m => m.Ncid == item.noCodeId).FirstOrDefault();

                        DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndDateTime);
                        DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartDateTime);

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
                            dbCheck.EndDateTime = endT;
                            db.SaveChanges();
                            #endregion

                            #region add new row in temp mode table which we are going to insert in mode table
                            Tbltcflossofentry tbltcflossofentry = new Tbltcflossofentry();
                            tbltcflossofentry.LossId = dbCheck.LossId;
                            tbltcflossofentry.MessageCodeId = dbCheck.MessageCodeId;
                            tbltcflossofentry.StartDateTime = endT;
                            tbltcflossofentry.EndDateTime = EndDateTime;
                            tbltcflossofentry.ReasonLevel1 = dbCheck.ReasonLevel1;
                            tbltcflossofentry.ReasonLevel2 = dbCheck.ReasonLevel2;
                            tbltcflossofentry.ReasonLevel3 = dbCheck.ReasonLevel3;
                            tbltcflossofentry.CorrectedDate = dbCheck.CorrectedDate;
                            tbltcflossofentry.MachineId = dbCheck.MachineId;
                            tbltcflossofentry.IsUpdate = dbCheck.IsUpdate;
                            tbltcflossofentry.IsArroval = dbCheck.IsArroval;
                            tbltcflossofentry.IsAccept = dbCheck.IsAccept;
                            tbltcflossofentry.NoOfReason = dbCheck.NoOfReason;
                            tbltcflossofentry.RejectReasonId = dbCheck.RejectReasonId;
                            tbltcflossofentry.ApprovalLevel = dbCheck.ApprovalLevel;
                            tbltcflossofentry.IsAccept1 = dbCheck.IsAccept1;
                            tbltcflossofentry.RejectReasonId1 = dbCheck.RejectReasonId1;
                            tbltcflossofentry.UpdateLevel = dbCheck.UpdateLevel;
                            db.Tbltcflossofentry.Add(tbltcflossofentry);
                            db.SaveChanges();
                            #endregion

                            #region Response assiging
                            NoCodeStartEndDateTime modeStartEndDateTime = new NoCodeStartEndDateTime();
                            modeStartEndDateTime.noCodeId = tbltcflossofentry.Ncid;
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
                        var tcflossdet = db.Tbltcflossofentry.Where(m => m.Ncid == uaWoId).FirstOrDefault();
                        if (tcflossdet != null)
                        {
                            int count = allIds.Count();
                            int id1 = Convert.ToInt32(allIds[i + 1]);
                            int lastid = Convert.ToInt32(allIds[count- 1]);
                            var tcfdet = db.Tbltcflossofentry.Where(m => m.Ncid == id1).FirstOrDefault();
                            var lasttcfrecord = db.Tbltcflossofentry.Where(m => m.Ncid == lastid).FirstOrDefault();
                           
                            DateTime lastetTime = Convert.ToDateTime(lasttcfrecord.EndDateTime);

                            DateTime tblfirststTime = Convert.ToDateTime(tcflossdet.StartDateTime);
                            DateTime tblEtTime = Convert.ToDateTime(tcflossdet.EndDateTime);
                            DateTime curEtTime = Convert.ToDateTime(data.endTime);
                            DateTime nextEtTime = Convert.ToDateTime(tcfdet.EndDateTime);
                            DateTime nextStTime = Convert.ToDateTime(tcfdet.StartDateTime);
                            DateTime PrevTime = Convert.ToDateTime(data.endTime).AddDays(-1);
                            DateTime NextTime = Convert.ToDateTime(data.endTime).AddDays(1);
                            if (curEtTime < tblfirststTime)
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
                                int durationInSecfirst = (int)curEtTime.Subtract(tblfirststTime).TotalSeconds;
                                int durationInSecSecond = (int)nextEtTime.Subtract(nextStTime).TotalSeconds;
                                if (durationInSecfirst > 120 && durationInSecSecond > 120)
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
                                        comobj.errorMsg = "You cannot able to split this duration, because splitting duration "+ curEtTime + " is equal to next record end time "+ nextEtTime + "";
                                    } 
                                    else if(curEtTime > nextEtTime)
                                    {
                                        comobj.isTure = false;
                                        comobj.errorMsg = "You cannot able to split this duration, because splitting duration " + curEtTime + " is Greater than next record end time " + nextEtTime + "";
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
                                            tcflossdet.EndDateTime = curEtTime;
                                            db.SaveChanges();

                                            if (tcfdet != null)
                                            {
                                                tcfdet.StartDateTime = curEtTime;
                                                tcfdet.LossId = tcflossdet.LossId;
                                                db.SaveChanges();
                                            }
                                            SplitDurationList objSplitDurationList = new SplitDurationList();
                                            objSplitDurationList = GetTheSpliDurationList(data.uaWOIdS);

                                            comobj.isTure = true;
                                            comobj.response = objSplitDurationList;
                                        }
                                    }
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "The Duration splited must be greater than 120 Seconds";
                                }
                                   
                            }
                            else if(PrevTime <=lastetTime && PrevTime >=tblfirststTime)
                            {
                                int durationInSecfirst = (int)PrevTime.Subtract(tblfirststTime).TotalSeconds;
                                int durationInSecSecond = (int)nextEtTime.Subtract(nextStTime).TotalSeconds;
                                if (durationInSecfirst > 120 && durationInSecSecond > 120)
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
                                    else if (curEtTime > nextEtTime)
                                    {
                                        comobj.isTure = false;
                                        comobj.errorMsg = "You cannot able to split this duration, because splitting duration " + curEtTime + " is Greater than next record end time " + nextEtTime + "";
                                    }
                                    else
                                    {
                                        tcflossdet.EndDateTime = PrevTime;
                                        db.SaveChanges();

                                        if (tcfdet != null)
                                        {
                                            tcfdet.StartDateTime = PrevTime;
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
                                    comobj.response = "The Duration splited must be greater than 120 Seconds";
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
                        var operatorMailId = db.TblEmployee.Where(m => m.EmpName == check.UserName).Select(m => m.EmailId).FirstOrDefault();

                        if (operatorMailId != null)
                        {
                            var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == data.machineId).Select(m => new { m.CellId, m.PlantId, m.ShopId }).FirstOrDefault();

                            string[] ids = data.tempModeIds.Split(',');
                            List<int> intArry = ids.ToList().Select(int.Parse).ToList();
                            var check1 = db.Tbltcflossofentry.Where(m => intArry.Contains(m.Ncid) && m.IsArroval == 1 && m.ApprovalLevel ==0).Select(m => new { m.IsArroval, m.ApprovalLevel }).GroupBy(m => new { m.IsArroval, m.ApprovalLevel }).ToList();
                            if (check1.Count > 0)
                            {
                                var toMail1 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 1 && m.IsDeleted==0).Select(m => m.FirstApproverToList).FirstOrDefault();

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

                            var check2 = db.Tbltcflossofentry.Where(m => intArry.Contains(m.Ncid) && m.ApprovalLevel == 1).Select(m => new { m.IsArroval, m.ApprovalLevel }).GroupBy(m => new { m.IsArroval, m.ApprovalLevel }).ToList();
                            if (check2.Count > 0)
                            {
                                var toMail2 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 1 && m.IsDeleted == 0).Select(m => m.SecondApproverToList).FirstOrDefault();
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

        #region UpdateOEETable

        //public bool CalculateOEEForYesterday(DateTime? StartDate, DateTime? EndDate)
        //{
        //    log.Error("CalculateOEEForYesterday" + StartDate + ":" + EndDate);
        //    bool ret = false;

        //    //MessageBox.Show("StartTime= " + StartDate + " EndTime= " + EndDate);

        //    DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1); 

        //    if (StartDate != null && EndDate != null)
        //    {
        //        fromdate = Convert.ToDateTime(StartDate);
        //        todate = Convert.ToDateTime(EndDate);
        //    }
        //    //fromdate = Convert.ToDateTime(DateTime.Now.ToString("2018-05-01"));
        //    //todate = Convert.ToDateTime(DateTime.Now.ToString("2018-10-31"));
        //    fromdate = StartDate ?? DateTime.Now.AddDays(-1);
        //    todate = EndDate ?? DateTime.Now.AddDays(-1);

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
        //            log.Error("MachineID" + macrow.MachineId);
        //            int MachineID = macrow.MachineId;

        //            try
        //            {
        //                var OEEDataPresent = db.Tbloeedashboardvariables.Where(m => m.Wcid == MachineID && m.StartDate == UsedDateForExcel).ToList();
        //                log.Error("OeeDataPresent" + OEEDataPresent.Count);
        //                if (OEEDataPresent.Count == 0)
        //                {
        //                    double green, red, yellow, blue, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
        //                    double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
        //                    double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;

        //                    MinorLosses = GetMinorLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                    log.Error("MinorLoss:" + MinorLosses);
        //                    if (MinorLosses < 0)
        //                    {
        //                        MinorLosses = 0;
        //                    }
        //                    blue = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                    log.Error("Blue:" + blue);
        //                    green = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
        //                    log.Error("green:" + green);
        //                    try
        //                    {
        //                        //Availability
        //                        SettingTime = GetSettingTime(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (SettingTime < 0)
        //                        {
        //                            SettingTime = 0;
        //                        }
        //                        ROALossess = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROA");
        //                        if (ROALossess < 0)
        //                        {
        //                            ROALossess = 0;
        //                        }
        //                        DownTimeBreakdown = GetDownTimeBreakdown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (DownTimeBreakdown < 0)
        //                        {
        //                            DownTimeBreakdown = 0;
        //                        }

        //                        //Performance
        //                        SummationOfSCTvsPP = GetSummationOfSCTvsPP(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (SummationOfSCTvsPP <= 0)
        //                        {
        //                            SummationOfSCTvsPP = 0;
        //                        }

        //                        //ROPLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROP");
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        log.Error("AVAlibility performance");
        //                    }

        //                    //Quality
        //                    try
        //                    {
        //                        ScrapQtyTime = GetScrapQtyTimeOfWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (ScrapQtyTime < 0)
        //                        {
        //                            ScrapQtyTime = 0;
        //                        }
        //                        ReWOTime = GetScrapQtyTimeOfRWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        if (ReWOTime < 0)
        //                        {
        //                            ReWOTime = 0;
        //                        }
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        log.Error("Quality");
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


        //                    //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //                    //{
        //                    var lossData = db.Tbllossofentry.Where(m => m.CorrectedDate == todayAsCorrectedDate && m.MachineId == MachineID).ToList();
        //                    log.Error("lossData" + lossData.Count);
        //                    foreach (var row in lossData)
        //                    {
        //                        int lossCodeID = Convert.ToInt32(row.MessageCodeId);
        //                        DateTime startDate = Convert.ToDateTime(row.StartDateTime);
        //                        DateTime endDate = Convert.ToDateTime(row.EndDateTime);
        //                        int duration = Convert.ToInt32(endDate.Subtract(startDate).TotalMinutes);

        //                        DataRow dr = DTLosses.Select("lossCodeID= '" + lossCodeID + "'").FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
        //                        if (dr != null)
        //                        {
        //                            int LossDurationPrev = Convert.ToInt32(dr["LossDuration"]); //get lossduration and update it.
        //                            dr["LossDuration"] = (LossDurationPrev + duration);
        //                        }
        //                        //}
        //                        else
        //                        {
        //                            DTLosses.Rows.Add(lossCodeID, duration);
        //                        }
        //                    }
        //                    //}
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
        //                    //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //                    //{
        //                    try
        //                    {
        //                        Tbloeedashboardvariables objoee = new Tbloeedashboardvariables();
        //                        objoee.PlantId = Convert.ToInt32(PlantIDS);
        //                        objoee.ShopId = Convert.ToInt32(ShopIDS);
        //                        objoee.CellId = Convert.ToInt32(CellIDS);
        //                        objoee.Wcid = Convert.ToInt32(MachineID);
        //                        objoee.StartDate = UsedDateForExcel;
        //                        objoee.EndDate = UsedDateForExcel;
        //                        objoee.MinorLosses = Math.Round(MinorLosses / 60, 2);
        //                        objoee.Blue = Math.Round(blue / 60, 2);
        //                        objoee.Green = Math.Round(green / 60, 2);
        //                        objoee.SettingTime = Math.Round(SettingTime, 2);
        //                        objoee.Roalossess = Math.Round(ROALossess / 60, 2);
        //                        objoee.DownTimeBreakdown = Math.Round(DownTimeBreakdown, 2);
        //                        objoee.SummationOfSctvsPp = Math.Round(SummationOfSCTvsPP, 2);
        //                        objoee.ScrapQtyTime = Math.Round(ScrapQtyTime, 2);
        //                        objoee.ReWotime = Math.Round(ReWOTime, 2);
        //                        objoee.Loss1Name = lossCode1;
        //                        objoee.Loss1Value = lossCodeVal1;
        //                        objoee.Loss2Name = lossCode2;
        //                        objoee.Loss2Value = lossCodeVal2;
        //                        objoee.Loss3Name = lossCode3;
        //                        objoee.Loss3Value = lossCodeVal3;
        //                        objoee.Loss4Name = lossCode4;
        //                        objoee.Loss4Value = lossCodeVal4;
        //                        objoee.Loss5Name = lossCode5;
        //                        objoee.Loss5Value = lossCodeVal5;
        //                        objoee.CreatedOn = DateTime.Now;
        //                        objoee.CreatedBy = 1;
        //                        objoee.IsDeleted = 0;
        //                        db.Tbloeedashboardvariables.Add(objoee);
        //                        db.SaveChanges();
        //                        log.Error("CalculateOEEForYesterday Insertion Completed");
        //                        ret = true;
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        log.Error("CalculateOEEForYesterday Insertion failed");
        //                        ret = false;
        //                    }
        //                    //finally
        //                    //{
        //                    //    mcInsertRows.close();
        //                    //}
        //                    //}
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                log.Error("CalculateOEEForYesterday MachineID not present");
        //                //IntoFile("MacID: " + MachineID + e.ToString());
        //            }
        //        }
        //        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
        //    }
        //    #endregion
        //    return ret;
        //}

        //public double GetMinorLosses(string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string dateString = currentdate.ToString("yyyy-MM-dd");

        //    log.Error("CorrectedDate:" + CorrectedDate);
        //    log.Error("MachineID:" + MachineID);
        //    log.Error("Colour:" + Colour);
        //    double minorloss = 0;
        //    try
        //    {
        //        #region commented
        //        //int count = 0;
        //        //var Data = db.tbldailyprodstatus.Where(m => m.IsDeleted == 0 && m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).OrderBy(m => m.StartTime).ToList();
        //        //foreach (var row in Data)
        //        //{
        //        //    if (row.ColorCode == "yellow")
        //        //    {
        //        //        count++;
        //        //    }
        //        //    else
        //        //    {
        //        //        if (count > 0 && count < 2)
        //        //        {
        //        //            minorloss += count;
        //        //            count = 0;

        //        //        }
        //        //        count = 0;
        //        //    }
        //        //}

        //        #endregion
        //        //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //        //{

        //        var MinorLossSummation = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == dateString && m.ColorCode == Colour && m.DurationInSec < 120 && m.IsCompleted == 1).Sum(m => m.DurationInSec);
        //        minorloss = Convert.ToDouble(MinorLossSummation);
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetMinorLosses" + ex);
        //        log.Error(ex.ToString());
        //    }
        //    return minorloss;
        //}
        //public double GetOPIDleBreakDown(string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");

        //    log.Error("CorrectedDate:" + CorrectedDate);
        //    log.Error("MachineID:" + MachineID);
        //    log.Error("Colour:" + Colour);

        //    double count = 0;

        //    try
        //    {
        //        //MsqlConnection mc = new MsqlConnection();
        //        //mc.open();
        //        ////operating
        //        //mc.open();
        //        //String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //        //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //        //DataTable OP = new DataTable();
        //        //da1.Fill(OP);
        //        //mc.close();
        //        //if (OP.Rows.Count != 0)
        //        //{
        //        //    count[0] = Convert.ToInt32(OP.Rows[0][0]);
        //        //}

        //        //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //        //{
        //        var blah = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.ColorCode == Colour).Sum(m => m.DurationInSec);
        //        count = Convert.ToDouble(blah);
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetOPIDleBreakDown" + ex);
        //        log.Error(ex.ToString());
        //    }
        //    return count;
        //}

        //public double GetSettingTime(string UsedDateForExcel, int MachineID)
        //{
        //    double settingTime = 0;
        //    try
        //    {

        //        int setupid = 0;
        //        string settingString = "Setup";
        //        var setupiddata = db.Tbllossescodes.Where(m => m.MessageType.Contains(settingString)).FirstOrDefault();
        //        if (setupiddata != null)
        //        {
        //            setupid = setupiddata.LossCodeId;
        //        }
        //        else
        //        {
        //            //Session["Error"] = "Unable to get Setup's ID";
        //            return -1;
        //        }
        //        // getting all setup's sublevels ids.
        //        //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //        //{
        //        var SettingIDs = db.Tbllossescodes.Where(m => m.LossCodesLevel1Id == setupid || m.LossCodesLevel2Id == setupid).Select(m => m.LossCodeId).ToList();


        //        //settingTime = (from row in db.tbllivelossofenties
        //        //where row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );


        //        var SettingData = db.Tbllivelossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
        //        foreach (var row in SettingData)
        //        {
        //            DateTime startTime = Convert.ToDateTime(row.StartDateTime);
        //            DateTime endTime = Convert.ToDateTime(row.EndDateTime);
        //            settingTime += endTime.Subtract(startTime).TotalMinutes;
        //        }
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetSettingTime" + ex);
        //    }
        //    return settingTime;
        //}
        //public double GetDownTimeLosses(string UsedDateForExcel, int MachineID, string contribute)
        //{
        //    double LossTime = 0;
        //    //string contribute = "ROA";
        //    // getting all ROA sublevels ids.Only those of IDLE.
        //    try
        //    {
        //        //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //        //{
        //        var SettingIDs = db.Tbllossescodes.Where(m => m.ContributeTo == contribute && (m.MessageType != "PM" || m.MessageType != "BREAKDOWN")).Select(m => m.LossCodeId).ToList();

        //        var SettingData = db.Tbllivelossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();

        //        var LossDuration = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.IsCompleted == 1 && m.DurationInSec > 120 && m.ColorCode == "YELLOW").Sum(m => m.DurationInSec);

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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetDownTimeLosses" + ex);
        //    }
        //    return LossTime;
        //}
        //public double GetDownTimeBreakdown(string UsedDateForExcel, int MachineID)
        //{
        //    if (MachineID == 18)
        //    {
        //    }
        //    double LossTime = 0;
        //    try
        //    {
        //        //using (i_facility_talContext dbLoss = new i_facility_talContext())
        //        //{
        //        var BreakdownData = db.Tblbreakdown.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetDownTimeBreakdown" + ex);
        //    }
        //    return LossTime;
        //}

        //public double GetSummationOfSCTvsPP(string UsedDateForExcel, int MachineID)
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
        //    try
        //    {
        //        //using (i_facility_talContext dbhmi = new i_facility_talContext())
        //        //{
        //        var PartsDataAll = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
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
        //                var multiWOData = db.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetSummationOfSCTvsPP" + ex);
        //    }
        //    return SummationofTime;
        //}

        ////Output in Seconds
        //public double GetGreen(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        //{
        //    double settingTime = 0;
        //    try
        //    {
        //        string stTime = StartTime.ToString("yyyy-MM-dd HH:mm:ss");
        //        //using (i_facility_talContext db = new i_facility_talContext())
        //        //{
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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetGreen" + ex);
        //    }
        //    return settingTime;
        //}

        //public double GetBlue(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        //{
        //    double settingTime = 0;
        //    try
        //    {
        //        //using (i_facility_talContext db = new i_facility_talContext())
        //        //{
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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetBlue" + ex);
        //    }
        //    return settingTime;
        //}

        //public double GetScrapQtyTimeOfWO(string UsedDateForExcel, int MachineID)
        //{
        //    double SQT = 0;
        //    try
        //    {
        //        //using (i_facility_talContext dbhmi = new i_facility_talContext())
        //        //{
        //        var PartsData = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 0).ToList();
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
        //            Double WODuration = GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

        //            if ((scrapQty + DeliveredQty) == 0)
        //            {
        //                SQT += 0;
        //            }
        //            else
        //            {
        //                SQT += ((WODuration / 60) / (scrapQty + DeliveredQty)) * scrapQty;
        //            }
        //        }
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetScrapQtyTimeOfWO" + ex);
        //    }
        //    return SQT;
        //}

        ////GOD
        //public double GetScrapQtyTimeOfRWO(string UsedDateForExcel, int MachineID)
        //{
        //    double SQT = 0;
        //    try
        //    {
        //        //using (i_facility_talContext dbhmi = new i_facility_talContext())
        //        //{
        //        var PartsData = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 1).ToList();
        //        foreach (var row in PartsData)
        //        {
        //            string partno = row.PartNo;
        //            string operationno = row.OperationNo;
        //            int scrapQty = Convert.ToInt32(row.RejQty);
        //            int DeliveredQty = Convert.ToInt32(row.DeliveredQty);
        //            DateTime startTime = Convert.ToDateTime(row.Date);
        //            DateTime endTime = Convert.ToDateTime(row.Time);
        //            Double WODuration = GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

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
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error("GetScrapQtyTimeOfRWO" + ex);
        //    }
        //    return SQT;
        //}



        #endregion

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

        //                                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tsal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + CorrectedDate + "' and MessageCodeID = '" + lossID + "' and DoneWithRow = 1  and "
        //                                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //                                    + " (  StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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

        //                                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tsal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + CorrectedDate + "' and MessageCodeID = '" + lossID + "' and DoneWithRow = 1  and "
        //                                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //                                    + " (  StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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
        //        String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tsal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and MessageCodeID IN ( " + SelfInspectionid + " ) and DoneWithRow = 1  and "
        //            + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //            + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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


        //        String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tsal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and DoneWithRow = 1  and "
        //            + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //            + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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

        //        String query1 = "SELECT StartTime,EndTime From [i_facility_tsal].[dbo].tblbreakdown WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and DoneWithRow = 1  and "
        //            + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //            + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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

        //        String query3 = "SELECT StartTime,EndTime,ModeID From [i_facility_tsal].[dbo].tblmode WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and ColorCode = 'yellow' and  DurationInSec < 120 and"
        //        + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //        + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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

        //        String query3 = "SELECT StartTime,EndTime,ModeID From [i_facility_tsal].[dbo].tblmode WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and ColorCode = 'green'  and"
        //           + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //           + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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

        //        String query3 = "SELECT StartTime,EndTime,ModeID From [i_facility_tsal].[dbo].tblmode WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and ColorCode = 'blue'  and"
        //           + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
        //           + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

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
