using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static DAS.EntityModels.CommonResponseWithMachineName;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.DAL
{
    public class SpliDurationDAL : ISplitDuration
    {

        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(SpliDurationDAL));
        public static IConfiguration configuration;

        public SpliDurationDAL(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }

        /// <summary>
        /// Get Plants
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetPlants()
        {
            CommonResponse1 obj = new CommonResponse1();

            var check = (from wf in db.Tblplant
                         where wf.IsDeleted == 0
                         select new
                         {
                             PlantId = wf.PlantId,
                             PlantName = wf.PlantName,
                             PlantDesc = wf.PlantDesc
                         }).ToList();
            if (check.Count > 0)
            {
                obj.isStatus = true;
                obj.response = check;
            }
            else
            {
                obj.isStatus = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        /// <summary>
        /// Get Shops
        /// </summary>
        /// <param name="plantId"></param>
        /// <returns></returns>
        public CommonResponse1 GetShops(int plantId)
        {
            CommonResponse1 obj = new CommonResponse1();
            var check = (from wf in db.Tblshop
                         where wf.IsDeleted == 0 && wf.PlantId == plantId
                         select new
                         {
                             ShopId = wf.ShopId,
                             ShopName = wf.ShopName,
                             ShopDesc = wf.ShopDesc
                         }).ToList();
            if (check.Count > 0)
            {
                obj.isStatus = true;
                obj.response = check;
            }
            else
            {
                obj.isStatus = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        /// <summary>
        /// Get Cells
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        public CommonResponse1 GetCells(int shopId)
        {
            CommonResponse1 obj = new CommonResponse1();
            var check = (from wf in db.Tblcell
                         where wf.IsDeleted == 0 && wf.ShopId == shopId
                         select new
                         {
                             CellId = wf.CellId,
                             CellName = wf.CellName,
                             CellDesc = wf.CellDesc
                         }).ToList();
            if (check.Count > 0)
            {
                obj.isStatus = true;
                obj.response = check;
            }
            else
            {
                obj.isStatus = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        /// <summary>
        /// Get Machines
        /// </summary>
        /// <param name="cellId"></param>
        /// <returns></returns>
        public CommonResponse1 GetMachines(int cellId)
        {
            CommonResponse1 obj = new CommonResponse1();
            var check = (from wf in db.Tblmachinedetails
                         where wf.CellId == cellId && wf.IsDeleted == 0
                         select new
                         {
                             MachineId = wf.MachineId,
                             MachineInvNo = wf.MachineInvNo,
                             MachineDispName = wf.MachineDispName,
                         }).ToList();
            if (check.Count > 0)
            {
                obj.isStatus = true;
                obj.response = check;
            }
            else
            {
                obj.isStatus = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        ///// <summary>
        ///// Get Mode Details
        ///// </summary>
        ///// <param name="MachineId"></param>
        ///// <returns></returns>
        //public CommonResponse1 GetModeDetails(int MachineId, string date)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        #region
        //        string CorrectedDate = "";
        //        try
        //        {
        //            string[] dt = date.Split('/');
        //            string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            CorrectedDate = frDate;
        //        }
        //        catch
        //        {
        //            CorrectedDate = date;
        //        }
        //        #endregion

        //        List<ModeStartEndDateTime> modeStartEndDateTimeList = new List<ModeStartEndDateTime>();

        //        #region get mode details either from temp table or mode table 

        //        var tempDataDet = db.TblTempMode.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate).ToList();

        //        if (tempDataDet.Count > 0)
        //        {
        //            foreach (var item in tempDataDet)
        //            {
        //                if (item.IsApproved == 0 || item.IsApproved == 3)
        //                {
        //                    db.TblTempMode.Remove(item);
        //                    db.SaveChanges();

        //                    var liveLoss = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == item.TempModeId).FirstOrDefault();
        //                    if (liveLoss != null)
        //                    {
        //                        db.TblTempLiveLossOfEntry.Remove(liveLoss);
        //                        db.SaveChanges();
        //                    }

        //                }
        //                else if (item.IsApproved == 1)
        //                {
        //                    ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
        //                    modeStartEndDateTime.message = "Already Send To Approval";
        //                    modeStartEndDateTimeList.Add(modeStartEndDateTime);
        //                }
        //            }
        //            var dbCheck = (from wf in db.Tblmode
        //                           where wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate && wf.Mode == "PowerOff"
        //                           select new
        //                           {
        //                               ModeId = wf.ModeId,
        //                               StartTime = wf.StartTime,
        //                               EndTime = wf.EndTime
        //                           }).ToList();

        //            if (dbCheck.Count > 0)
        //            {
        //                foreach (var item1 in dbCheck)
        //                {
        //                    ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
        //                    DateTime startDateTime = Convert.ToDateTime(item1.StartTime);
        //                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    string[] ids = dt.Split();
        //                    modeStartEndDateTime.startDate = ids[0];
        //                    modeStartEndDateTime.startTime = ids[1];
        //                    DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
        //                    string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    string[] ids1 = dt1.Split();
        //                    modeStartEndDateTime.endDate = ids1[0];
        //                    modeStartEndDateTime.endTime = ids1[1];
        //                    modeStartEndDateTime.modeId = item1.ModeId;
        //                    modeStartEndDateTimeList.Add(modeStartEndDateTime);
        //                    obj.isStatus = true;
        //                    obj.response = modeStartEndDateTimeList.OrderBy(m => m.startTime);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            var dbCheck = (from wf in db.Tblmode
        //                           where wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate && wf.Mode == "PowerOff"
        //                           select new
        //                           {
        //                               ModeId = wf.ModeId,
        //                               StartTime = wf.StartTime,
        //                               EndTime = wf.EndTime
        //                           }).ToList();

        //            if (dbCheck.Count > 0)
        //            {
        //                foreach (var item1 in dbCheck)
        //                {
        //                    ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
        //                    DateTime startDateTime = Convert.ToDateTime(item1.StartTime);
        //                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    string[] ids = dt.Split();
        //                    modeStartEndDateTime.startDate = ids[0];
        //                    modeStartEndDateTime.startTime = ids[1];
        //                    DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
        //                    string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    string[] ids1 = dt1.Split();
        //                    modeStartEndDateTime.endDate = ids1[0];
        //                    modeStartEndDateTime.endTime = ids1[1];
        //                    modeStartEndDateTime.modeId = item1.ModeId;
        //                    modeStartEndDateTimeList.Add(modeStartEndDateTime);
        //                    obj.isStatus = true;
        //                    obj.response = modeStartEndDateTimeList.OrderBy(m => m.startTime);
        //                }
        //            }
        //        }
        //        #endregion
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        /// <summary>
        /// Validate End time and send start date & start Time
        /// </summary>
        /// <param name="ModeId"></param>
        /// <param name="EndTime"></param>
        /// <returns></returns>
        public CommonResponse1 ValidateEndtime(int ModeId, string EndTime)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.Tblmode.Where(m => m.ModeId == ModeId).FirstOrDefault();
                if (check != null)
                {
                    DateTime StartDateTime = Convert.ToDateTime(check.StartTime);
                    string dt = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids = dt.Split();
                    string startDate = ids[0];
                    string startTime = ids[1];
                    DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
                    string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids1 = dt1.Split();
                    string endDate = ids1[0];
                    string endTime = ids1[1];
                    string endTimeLast = ids1[0] + " " + EndTime;
                    DateTime endT = Convert.ToDateTime(endTimeLast);
                    if (endT < EndDateTime && endT > StartDateTime)
                    {
                        obj.isStatus = true;
                        obj.response = "Valid End Time";
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "Please Enter the EndTime with in this " + endTime;
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

        /// <summary>
        /// Add mode details To Temp Table before saving it to mode table
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 AddModeDetailsToTempTable(List<ModeStartEndDateTime> data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                int i = 0;
                List<ModeStartEndDateTime> modeStartEndDateTimeList = new List<ModeStartEndDateTime>();
                foreach (var item in data)
                {
                    i++;
                    //modeStartEndDateTimeList.Add(item);
                    if (data.Count == i)
                    {
                        var dbCheck = db.TblTempMode.Where(m => m.TempModeId == item.tempModeId).FirstOrDefault();

                        DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                        DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartTime);

                        string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids = dt.Split();
                        string endDate = ids[0];
                        string endTime = ids[1];

                        string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids1 = dt1.Split();
                        string startDate = ids1[0];
                        string startTime = ids1[1];

                        string endTimeLast = ids1[0] + " " + item.endTime;
                        DateTime endT = Convert.ToDateTime(endTimeLast);

                        #region jugad
                        EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                        StartDateTime = Convert.ToDateTime(dbCheck.StartTime);
                        bool flag = false;
                        if ((EndDateTime.Date - StartDateTime.Date).Days > 0)
                        {
                            flag = true;
                        }
                        #endregion

                        if (endT < EndDateTime && endT > StartDateTime)
                        {
                            #region Update old row
                            dbCheck.EndTime = endT;
                            dbCheck.Mode = "PowerOff";
                            var durationInSec = (endT - StartDateTime).TotalSeconds;
                            dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                            db.SaveChanges();

                            ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                            modeStartEndDateTime.tempModeId = dbCheck.TempModeId;
                            modeStartEndDateTime.startDate = startDate;
                            modeStartEndDateTime.startTime = startTime;
                            modeStartEndDateTime.endDate = startDate;
                            modeStartEndDateTime.endTime = item.endTime;
                            modeStartEndDateTime.modeId = dbCheck.ModeId;
                            modeStartEndDateTimeList.Add(modeStartEndDateTime);
                            #endregion

                            #region add new row in temp mode table which we are going to insert in mode table
                            TblTempMode tblTempMode = new TblTempMode();
                            tblTempMode.MachineId = dbCheck.MachineId;
                            tblTempMode.ModeId = dbCheck.ModeId;
                            tblTempMode.Mode = "PowerOff";
                            tblTempMode.ColorCode = "blue";
                            tblTempMode.InsertedOn = DateTime.Now;
                            tblTempMode.InsertedBy = 1;
                            tblTempMode.IsSaved = 0;
                            tblTempMode.CorrectedDate = dbCheck.CorrectedDate;
                            tblTempMode.IsDeleted = 0;
                            tblTempMode.StartTime = endT;
                            tblTempMode.EndTime = EndDateTime;
                            tblTempMode.IsCompleted = 1;
                            tblTempMode.IsApproved = 0;
                            var durationInSec11 = (EndDateTime - endT).TotalSeconds;
                            tblTempMode.DurationInSec = Convert.ToInt32(durationInSec11);
                            db.TblTempMode.Add(tblTempMode);
                            db.SaveChanges();
                            #endregion

                            #region Response assiging
                            ModeStartEndDateTime modeStartEndDateTime1 = new ModeStartEndDateTime();
                            modeStartEndDateTime1.tempModeId = tblTempMode.TempModeId;
                            modeStartEndDateTime1.startDate = startDate;
                            modeStartEndDateTime1.startTime = item.endTime;
                            modeStartEndDateTime1.endDate = endDate;
                            modeStartEndDateTime1.endTime = endTime;
                            modeStartEndDateTime1.modeId = tblTempMode.ModeId;
                            modeStartEndDateTimeList.Add(modeStartEndDateTime1);
                            //obj.isStatus = true;
                            //obj.response = modeStartEndDateTimeList;
                            #endregion 

                            var result = db.TblTempMode.Where(p => p.ModeId == tblTempMode.ModeId && p.CorrectedDate == tblTempMode.CorrectedDate && p.MachineId == tblTempMode.MachineId).ToList();

                            var result1 = result.Where(p => modeStartEndDateTimeList.All(p2 => p2.tempModeId != p.TempModeId)).ToList();
                            foreach (var tempData in result1)
                            {
                                DateTime EndDateTime1 = Convert.ToDateTime(tempData.EndTime);
                                DateTime StartDateTime1 = Convert.ToDateTime(tempData.StartTime);

                                string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids2 = dt2.Split();
                                string endDate2 = ids2[0];
                                string endTime2 = ids2[1];

                                string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids3 = dt3.Split();
                                string startDate3 = ids3[0];
                                string startTime3 = ids3[1];

                                ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                                modeStartEndDateTime2.tempModeId = tempData.TempModeId;
                                modeStartEndDateTime2.startDate = startDate3;
                                modeStartEndDateTime2.startTime = startTime3;
                                modeStartEndDateTime2.endDate = endDate2;
                                modeStartEndDateTime2.endTime = endTime2;
                                modeStartEndDateTime2.modeId = tempData.ModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime2);

                            }
                        }
                        else if (flag == true)
                        {
                            string tempStartDate = StartDateTime.ToString("yyyy-MM-dd");
                            string tempEndDate = EndDateTime.ToString("yyyy-MM-dd");

                            DateTime endTT = Convert.ToDateTime(tempEndDate + " " + item.endTime);
                            string[] endTTString = EndDateTime.ToString().Split(' ');
                            DateTime endTT1 = Convert.ToDateTime(tempEndDate + " " + endTTString[1]);

                            if (endTT < EndDateTime && endTT > StartDateTime)
                            {
                                #region Update old row
                                dbCheck.EndTime = endTT;
                                dbCheck.Mode = "PowerOff";
                                var durationInSec = (endTT - StartDateTime).TotalSeconds;
                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                                db.SaveChanges();

                                ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                                modeStartEndDateTime.tempModeId = dbCheck.TempModeId;
                                modeStartEndDateTime.startDate = startDate;
                                modeStartEndDateTime.startTime = startTime;
                                modeStartEndDateTime.endDate = tempEndDate;
                                modeStartEndDateTime.endTime = item.endTime;
                                modeStartEndDateTime.modeId = dbCheck.ModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime);
                                #endregion

                                #region add new row in temp mode table which we are going to insert in mode table
                                TblTempMode tblTempMode = new TblTempMode();
                                tblTempMode.MachineId = dbCheck.MachineId;
                                tblTempMode.ModeId = dbCheck.ModeId;
                                tblTempMode.Mode = "PowerOff";
                                tblTempMode.ColorCode = "blue";
                                tblTempMode.InsertedOn = DateTime.Now;
                                tblTempMode.InsertedBy = 1;
                                tblTempMode.IsSaved = 0;
                                tblTempMode.CorrectedDate = dbCheck.CorrectedDate;
                                tblTempMode.IsDeleted = 0;
                                tblTempMode.StartTime = endTT;
                                tblTempMode.EndTime = endTT1;
                                tblTempMode.IsCompleted = 1;
                                tblTempMode.IsApproved = 0;
                                var durationInSec11 = (endTT1 - endT).TotalSeconds;
                                tblTempMode.DurationInSec = Convert.ToInt32(durationInSec11);
                                db.TblTempMode.Add(tblTempMode);
                                db.SaveChanges();
                                #endregion

                                #region Response assiging
                                ModeStartEndDateTime modeStartEndDateTime1 = new ModeStartEndDateTime();
                                modeStartEndDateTime1.tempModeId = tblTempMode.TempModeId;
                                modeStartEndDateTime1.startDate = tempEndDate;
                                modeStartEndDateTime1.startTime = item.endTime;
                                modeStartEndDateTime1.endDate = tempEndDate;
                                modeStartEndDateTime1.endTime = endTTString[1];
                                modeStartEndDateTime1.modeId = tblTempMode.ModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime1);

                                #endregion

                                var result = db.TblTempMode.Where(p => p.ModeId == tblTempMode.ModeId && p.CorrectedDate == tblTempMode.CorrectedDate && p.MachineId == tblTempMode.MachineId).ToList();

                                var result1 = result.Where(p => modeStartEndDateTimeList.All(p2 => p2.tempModeId != p.TempModeId)).ToList();
                                foreach (var tempData in result1)
                                {
                                    DateTime EndDateTime1 = Convert.ToDateTime(tempData.EndTime);
                                    DateTime StartDateTime1 = Convert.ToDateTime(tempData.StartTime);

                                    string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                    string[] ids2 = dt2.Split();
                                    string endDate2 = ids2[0];
                                    string endTime2 = ids2[1];

                                    string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                    string[] ids3 = dt3.Split();
                                    string startDate3 = ids3[0];
                                    string startTime3 = ids3[1];

                                    ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                                    modeStartEndDateTime2.tempModeId = tempData.TempModeId;
                                    modeStartEndDateTime2.startDate = startDate3;
                                    modeStartEndDateTime2.startTime = startTime3;
                                    modeStartEndDateTime2.endDate = endDate2;
                                    modeStartEndDateTime2.endTime = endTime2;
                                    modeStartEndDateTime2.modeId = tempData.ModeId;
                                    modeStartEndDateTimeList.Add(modeStartEndDateTime2);

                                }
                            }
                            else
                            {
                                var result = db.TblTempMode.Where(p => p.ModeId == dbCheck.ModeId && p.CorrectedDate == dbCheck.CorrectedDate && p.MachineId == dbCheck.MachineId).ToList();

                                //var result1 = result.Where(p => modeStartEndDateTimeList.All(p2 => p2.tempModeId != p.TempModeId)).ToList();
                                foreach (var tempData in result)
                                {
                                    DateTime EndDateTime1 = Convert.ToDateTime(tempData.EndTime);
                                    DateTime StartDateTime1 = Convert.ToDateTime(tempData.StartTime);

                                    string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                    string[] ids2 = dt2.Split();
                                    string endDate2 = ids2[0];
                                    string endTime2 = ids2[1];

                                    string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                    string[] ids3 = dt3.Split();
                                    string startDate3 = ids3[0];
                                    string startTime3 = ids3[1];

                                    ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                                    modeStartEndDateTime2.tempModeId = tempData.TempModeId;
                                    modeStartEndDateTime2.startDate = startDate3;
                                    modeStartEndDateTime2.startTime = startTime3;
                                    modeStartEndDateTime2.endDate = endDate2;
                                    modeStartEndDateTime2.endTime = endTime2;
                                    modeStartEndDateTime2.modeId = tempData.ModeId;
                                    modeStartEndDateTimeList.Add(modeStartEndDateTime2);

                                }
                            }
                        }
                        else
                        {
                            var result = db.TblTempMode.Where(p => p.ModeId == dbCheck.ModeId && p.CorrectedDate == dbCheck.CorrectedDate && p.MachineId == dbCheck.MachineId).ToList();

                            //var result1 = result.Where(p => modeStartEndDateTimeList.All(p2 => p2.tempModeId != p.TempModeId)).ToList();
                            foreach (var tempData in result)
                            {
                                DateTime EndDateTime1 = Convert.ToDateTime(tempData.EndTime);
                                DateTime StartDateTime1 = Convert.ToDateTime(tempData.StartTime);

                                string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids2 = dt2.Split();
                                string endDate2 = ids2[0];
                                string endTime2 = ids2[1];

                                string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids3 = dt3.Split();
                                string startDate3 = ids3[0];
                                string startTime3 = ids3[1];

                                ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                                modeStartEndDateTime2.tempModeId = tempData.TempModeId;
                                modeStartEndDateTime2.startDate = startDate3;
                                modeStartEndDateTime2.startTime = startTime3;
                                modeStartEndDateTime2.endDate = endDate2;
                                modeStartEndDateTime2.endTime = endTime2;
                                modeStartEndDateTime2.modeId = tempData.ModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime2);
                            }
                        }
                    }
                }

                modeStartEndDateTimeList = modeStartEndDateTimeList.GroupBy(e => new { e.endDate, e.endTime, e.message, e.modeId, e.startDate, e.startTime, e.tempModeId }).Select(g => g.First()).OrderBy(m => m.tempModeId).ToList();
                obj.isStatus = true;
                obj.response = modeStartEndDateTimeList;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// After clicking save button it should save in the main mode table
        /// </summary>
        /// <param name="tempModeIds"></param>
        /// <returns></returns>
        public CommonResponse1 ApprovedSendToMainModeDetails(string tempModeIds)
        {
            CommonResponse1 obj = new CommonResponse1();
            bool result = false;
            try
            {
                string[] ids = tempModeIds.Split(',');
                int machineID = 0;
                string correctedDate = "";
                foreach (var item in ids)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsApproved == 1).FirstOrDefault();
                    var dbCheck1 = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsApproved == 2 && m.ApproveLevel == 1).FirstOrDefault();

                    if (check != null)
                    {
                        machineID = check.MachineId;
                        correctedDate = check.CorrectedDate;

                        #region update approved mode ids
                        check.IsApproved = 2;
                        check.ApproveLevel = 1;
                        check.IsHold = 0;
                        check.ApprovedDate = DateTime.Now;
                        db.SaveChanges();
                        #endregion

                        #region update in tbl mode
                        var dbCheck = db.Tblmode.Where(m => m.ModeId == check.ModeId).FirstOrDefault();
                        if (dbCheck != null)
                        {
                            db.Tblmode.Remove(dbCheck);
                            db.SaveChanges();
                        }
                        #endregion

                        var getMachineDetails = (from wf in db.Tblmachinedetails
                                                 where wf.MachineId == machineID && wf.IsDeleted == 0
                                                 select new
                                                 {
                                                     PlantId = wf.PlantId,
                                                     ShopId = wf.ShopId,
                                                     CellId = wf.CellId
                                                 }).FirstOrDefault();

                        var mailInfo = (from wf in db.TblTcfApprovedMaster
                                        where wf.TcfModuleId == 3 && (wf.CellId == getMachineDetails.CellId || wf.ShopId == getMachineDetails.ShopId) && wf.SecondApproverToList != null && wf.SecondApproverToList != ""
                                        && wf.SecondApproverCcList != null && wf.SecondApproverCcList != ""
                                        select new
                                        {
                                            SecondApproverToList = wf.SecondApproverToList,
                                            SecondApproverCcList = wf.SecondApproverCcList
                                        }).ToList();

                        if (mailInfo.Count == 0)
                        {
                            #region add new row in live mode table which we are going to insert in mode table
                            Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                            tbllivemodedb.MachineId = check.MachineId;
                            tbllivemodedb.Mode = check.Mode;
                            tbllivemodedb.ColorCode = check.ColorCode;
                            tbllivemodedb.InsertedOn = DateTime.Now;
                            tbllivemodedb.InsertedBy = 1;
                            tbllivemodedb.CorrectedDate = check.CorrectedDate;
                            tbllivemodedb.IsDeleted = 0;
                            tbllivemodedb.StartTime = check.StartTime;
                            tbllivemodedb.EndTime = check.EndTime;
                            tbllivemodedb.IsCompleted = 1;
                            tbllivemodedb.DurationInSec = check.DurationInSec;
                            db.Tbllivemodedb.Add(tbllivemodedb);
                            db.SaveChanges();
                            #endregion

                            #region add new row in mode table which we have inserted in live mode table in last region
                            Tblmode tblmode = new Tblmode();
                            tblmode.MachineId = check.MachineId;
                            tblmode.ColorCode = tbllivemodedb.ColorCode;
                            tblmode.ModeId = tbllivemodedb.ModeId;
                            tblmode.Mode = tbllivemodedb.Mode;
                            tblmode.IsDeleted = 0;
                            tblmode.InsertedOn = DateTime.Now;
                            tblmode.InsertedBy = 1;
                            tblmode.CorrectedDate = check.CorrectedDate;
                            tblmode.StartTime = tbllivemodedb.StartTime;
                            tblmode.EndTime = tbllivemodedb.EndTime;
                            tblmode.IsCompleted = 1;
                            tblmode.DurationInSec = tbllivemodedb.DurationInSec;
                            db.Tblmode.Add(tblmode);
                            db.SaveChanges();
                            #endregion

                            #region insert loss details in live loss of entry

                            var liveLoss = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                            if (liveLoss != null)
                            {
                                var checkBr = db.TblTempMode.Where(m => m.TempModeId == tempModeId).Select(m => m.Mode).FirstOrDefault();
                                if (checkBr == "IDLE")
                                {
                                    int checkDuration = Convert.ToInt32(db.TblTempMode.Where(m => m.TempModeId == tempModeId).Select(m => m.DurationInSec).FirstOrDefault());
                                    if (checkDuration > 120)
                                    {
                                        Tbllivelossofentry tbllivelossofentry = new Tbllivelossofentry();
                                        tbllivelossofentry.StartDateTime = liveLoss.StartDateTime;
                                        tbllivelossofentry.EntryTime = liveLoss.EntryTime;
                                        tbllivelossofentry.EndDateTime = liveLoss.EndDateTime;
                                        tbllivelossofentry.CorrectedDate = liveLoss.CorrectedDate;
                                        tbllivelossofentry.MachineId = liveLoss.MachineId;
                                        tbllivelossofentry.Shift = liveLoss.Shift;
                                        tbllivelossofentry.MessageCodeId = Convert.ToInt32(liveLoss.MessageCodeId);
                                        tbllivelossofentry.MessageCode = liveLoss.MessageCode;
                                        tbllivelossofentry.MessageDesc = liveLoss.MessageDesc;
                                        tbllivelossofentry.IsUpdate = 1;
                                        tbllivelossofentry.DoneWithRow = 1;
                                        tbllivelossofentry.IsStart = 0;
                                        tbllivelossofentry.IsScreen = 0;
                                        tbllivelossofentry.ForRefresh = 0;
                                        db.Tbllivelossofentry.Add(tbllivelossofentry);
                                        db.SaveChanges();
                                        obj.isStatus = true;
                                        obj.response = "Added Successfully";

                                        Tbllossofentry tbllossofentry = new Tbllossofentry();
                                        tbllossofentry.LossId = tbllivelossofentry.LossId;
                                        tbllossofentry.MachineId = tbllivelossofentry.MachineId;
                                        tbllossofentry.StartDateTime = tbllivelossofentry.StartDateTime;
                                        tbllossofentry.EntryTime = tbllivelossofentry.EntryTime;
                                        tbllossofentry.EndDateTime = tbllivelossofentry.EndDateTime;
                                        tbllossofentry.MessageCodeId = tbllivelossofentry.MessageCodeId;
                                        tbllossofentry.MessageCode = tbllivelossofentry.MessageCode;
                                        tbllossofentry.MessageDesc = tbllivelossofentry.MessageDesc;
                                        tbllossofentry.CorrectedDate = tbllivelossofentry.CorrectedDate;
                                        tbllossofentry.IsUpdate = 1;
                                        tbllossofentry.DoneWithRow = 1;
                                        tbllossofentry.IsStart = 0;
                                        tbllossofentry.IsScreen = 0;
                                        tbllossofentry.ForRefresh = 0;
                                        db.Tbllossofentry.Add(tbllossofentry);
                                        db.SaveChanges();
                                        obj.isStatus = true;
                                        obj.response = "Added Successfully";
                                    }
                                }
                                else if (checkBr == "BREAKDOWN")
                                {
                                    Tblbreakdown tblbreakdown = new Tblbreakdown();
                                    tblbreakdown.StartTime = liveLoss.StartDateTime;
                                    tblbreakdown.EndTime = liveLoss.EndDateTime;
                                    tblbreakdown.BreakDownCode = liveLoss.MessageCodeId;
                                    tblbreakdown.MachineId = liveLoss.MachineId;
                                    tblbreakdown.CorrectedDate = liveLoss.CorrectedDate;
                                    tblbreakdown.Shift = liveLoss.Shift;
                                    tblbreakdown.MessageDesc = liveLoss.MessageDesc;
                                    tblbreakdown.MessageCode = liveLoss.MessageCode;
                                    tblbreakdown.DoneWithRow = 1;
                                    db.Tblbreakdown.Add(tblbreakdown);
                                    db.SaveChanges();
                                }
                            }
                            #endregion
                        }
                    }
                    else if (dbCheck1 != null)
                    {
                        machineID = dbCheck1.MachineId;
                        correctedDate = dbCheck1.CorrectedDate;
                        #region update approved mode ids
                        dbCheck1.IsApproved = 2;
                        dbCheck1.ApproveLevel = 2;
                        dbCheck1.IsHold = 0;
                        dbCheck1.ApprovedDate = DateTime.Now;
                        db.SaveChanges();
                        #endregion

                        #region update in tbl mode
                        var dbCheck = db.Tblmode.Where(m => m.ModeId == dbCheck1.ModeId).FirstOrDefault();
                        if (dbCheck != null)
                        {
                            db.Tblmode.Remove(dbCheck);
                            db.SaveChanges();
                        }
                        #endregion

                        #region add new row in live mode table which we are going to insert in mode table
                        Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                        tbllivemodedb.MachineId = dbCheck1.MachineId;
                        tbllivemodedb.Mode = dbCheck1.Mode;
                        tbllivemodedb.ColorCode = dbCheck1.ColorCode;
                        tbllivemodedb.InsertedOn = DateTime.Now;
                        tbllivemodedb.InsertedBy = 1;
                        tbllivemodedb.CorrectedDate = dbCheck1.CorrectedDate;
                        tbllivemodedb.IsDeleted = 0;
                        tbllivemodedb.StartTime = dbCheck1.StartTime;
                        tbllivemodedb.EndTime = dbCheck1.EndTime;
                        tbllivemodedb.IsCompleted = 1;
                        tbllivemodedb.DurationInSec = dbCheck1.DurationInSec;
                        db.Tbllivemodedb.Add(tbllivemodedb);
                        db.SaveChanges();
                        #endregion

                        #region add new row in mode table which we have inserted in live mode table in last region
                        Tblmode tblmode = new Tblmode();
                        tblmode.MachineId = dbCheck1.MachineId;
                        tblmode.ColorCode = tbllivemodedb.ColorCode;
                        tblmode.ModeId = tbllivemodedb.ModeId;
                        tblmode.Mode = tbllivemodedb.Mode;
                        tblmode.IsDeleted = 0;
                        tblmode.InsertedOn = DateTime.Now;
                        tblmode.InsertedBy = 1;
                        tblmode.CorrectedDate = dbCheck1.CorrectedDate;
                        tblmode.StartTime = tbllivemodedb.StartTime;
                        tblmode.EndTime = tbllivemodedb.EndTime;
                        tblmode.IsCompleted = 1;
                        tblmode.DurationInSec = tbllivemodedb.DurationInSec;
                        db.Tblmode.Add(tblmode);
                        db.SaveChanges();
                        #endregion

                        #region insert loss details in live loss of entry

                        var liveLoss = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                        if (liveLoss != null)
                        {
                            var checkBr = db.TblTempMode.Where(m => m.TempModeId == tempModeId).Select(m => m.Mode).FirstOrDefault();
                            if (checkBr == "IDLE")
                            {
                                int checkDuration = Convert.ToInt32(db.TblTempMode.Where(m => m.TempModeId == tempModeId).Select(m => m.DurationInSec).FirstOrDefault());
                                if (checkDuration > 120)
                                {
                                    Tbllivelossofentry tbllivelossofentry = new Tbllivelossofentry();
                                    tbllivelossofentry.StartDateTime = liveLoss.StartDateTime;
                                    tbllivelossofentry.EntryTime = liveLoss.EntryTime;
                                    tbllivelossofentry.EndDateTime = liveLoss.EndDateTime;
                                    tbllivelossofentry.CorrectedDate = liveLoss.CorrectedDate;
                                    tbllivelossofentry.MachineId = liveLoss.MachineId;
                                    tbllivelossofentry.Shift = liveLoss.Shift;
                                    tbllivelossofentry.MessageCodeId = Convert.ToInt32(liveLoss.MessageCodeId);
                                    tbllivelossofentry.MessageCode = liveLoss.MessageCode;
                                    tbllivelossofentry.MessageDesc = liveLoss.MessageDesc;
                                    tbllivelossofentry.IsUpdate = 1;
                                    tbllivelossofentry.DoneWithRow = 1;
                                    tbllivelossofentry.IsStart = 0;
                                    tbllivelossofentry.IsScreen = 0;
                                    tbllivelossofentry.ForRefresh = 0;
                                    db.Tbllivelossofentry.Add(tbllivelossofentry);
                                    db.SaveChanges();
                                    obj.isStatus = true;
                                    obj.response = "Added Successfully";

                                    Tbllossofentry tbllossofentry = new Tbllossofentry();
                                    tbllossofentry.LossId = tbllivelossofentry.LossId;
                                    tbllossofentry.MachineId = tbllivelossofentry.MachineId;
                                    tbllossofentry.StartDateTime = tbllivelossofentry.StartDateTime;
                                    tbllossofentry.EntryTime = tbllivelossofentry.EntryTime;
                                    tbllossofentry.EndDateTime = tbllivelossofentry.EndDateTime;
                                    tbllossofentry.MessageCodeId = tbllivelossofentry.MessageCodeId;
                                    tbllossofentry.MessageCode = tbllivelossofentry.MessageCode;
                                    tbllossofentry.MessageDesc = tbllivelossofentry.MessageDesc;
                                    tbllossofentry.CorrectedDate = tbllivelossofentry.CorrectedDate;
                                    tbllossofentry.IsUpdate = 1;
                                    tbllossofentry.DoneWithRow = 1;
                                    tbllossofentry.IsStart = 0;
                                    tbllossofentry.IsScreen = 0;
                                    tbllossofentry.ForRefresh = 0;
                                    db.Tbllossofentry.Add(tbllossofentry);
                                    db.SaveChanges();
                                    obj.isStatus = true;
                                    obj.response = "Added Successfully";
                                }
                            }
                            else if (checkBr == "BREAKDOWN")
                            {
                                Tblbreakdown tblbreakdown = new Tblbreakdown();
                                tblbreakdown.StartTime = liveLoss.StartDateTime;
                                tblbreakdown.EndTime = liveLoss.EndDateTime;
                                tblbreakdown.BreakDownCode = liveLoss.MessageCodeId;
                                tblbreakdown.MachineId = liveLoss.MachineId;
                                tblbreakdown.CorrectedDate = liveLoss.CorrectedDate;
                                tblbreakdown.Shift = liveLoss.Shift;
                                tblbreakdown.MessageDesc = liveLoss.MessageDesc;
                                tblbreakdown.MessageCode = liveLoss.MessageCode;
                                tblbreakdown.DoneWithRow = 1;
                                db.Tblbreakdown.Add(tblbreakdown);
                                db.SaveChanges();
                            }
                            #endregion
                        }
                    }
                }
                #region call update data in report method

                var checkData2 = db.TblTempMode.Where(m => m.MachineId == machineID && m.ApproveLevel == 1 && m.CorrectedDate == correctedDate).ToList();

                var getMacDetails = (from wf in db.Tblmachinedetails
                                     where wf.MachineId == machineID && wf.IsDeleted == 0
                                     select new
                                     {
                                         PlantId = wf.PlantId,
                                         ShopId = wf.ShopId,
                                         CellId = wf.CellId
                                     }).FirstOrDefault();

                var mailInfo1 = (from wf in db.TblTcfApprovedMaster
                                 where wf.TcfModuleId == 3 && (wf.CellId == getMacDetails.CellId || wf.ShopId == getMacDetails.ShopId) && wf.SecondApproverToList != null && wf.SecondApproverToList != ""
                                 && wf.SecondApproverCcList != null && wf.SecondApproverCcList != ""
                                 select new
                                 {
                                     SecondApproverToList = wf.SecondApproverToList,
                                     SecondApproverCcList = wf.SecondApproverCcList
                                 }).ToList();

                var checkData3 = db.TblTempMode.Where(m => m.MachineId == machineID && m.ApproveLevel == 2 && m.CorrectedDate == correctedDate).ToList();

                if (checkData2.Count > 0)
                {
                    if (mailInfo1.Count == 0)
                    {
                        #region mimics updation

                        UpdateMimicsDetails(machineID, correctedDate);

                        #endregion

                        UpdateInReport(machineID, correctedDate);

                        #region Sending last Level mail details

                        int sl = 1;
                        string MachineName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();
                        string PlantName = db.Tblplant.Where(m => m.PlantId == getMacDetails.PlantId).Select(m => m.PlantName).FirstOrDefault();
                        string ShopName = db.Tblshop.Where(m => m.ShopId == getMacDetails.ShopId).Select(m => m.ShopName).FirstOrDefault();
                        string cellName = db.Tblcell.Where(m => m.CellId == getMacDetails.CellId).Select(m => m.CellName).FirstOrDefault();
                        int machineId1 = machineID;

                        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\DataLossTcfTemplate1.html");
                        string htmlStr = File.ReadAllText(reader);


                        //string[] ids1 = tempModeIds.Split(',');

                        string mainTemp = "<tr><td><span>{{slno}}</span> </td> <td><span>{{machinename}}</span> </td><td><span>{{mode}}" +
                            "</span> </td><td><span>{{rl1}}</span></td><td><span>{{rl2}}</span></td><td><span>{{rl3}}</span></  td><td><span>{{startTime}}</span></td><td><span>{{endTime}}</span></td></tr>";
                        string strReplace = "";

                        foreach (var item1 in checkData2)
                        {
                            var check1 = db.TblTempMode.Where(m => m.TempModeId == item1.TempModeId && m.IsApproved == 2 && m.ApproveLevel == 2).FirstOrDefault();

                            var lossData = (from wf in db.TblTempLiveLossOfEntry
                                            join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                            where wf.TempModeId == item1.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == machineID && wf.CorrectedDate == correctedDate
                                            select new
                                            {
                                                LossCodeId = lc.LossCodeId,
                                                LossCode = lc.LossCode,
                                                LossCodesLevel = lc.LossCodesLevel,
                                                LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                StartDateTime = wf.StartDateTime,
                                                EndDateTime = wf.EndDateTime,
                                                MachineId = wf.MachineId,
                                                MessageCodeId = wf.MessageCodeId,
                                                CorrectedDate = wf.CorrectedDate,
                                                TempLossId = wf.TempLossId,

                                                ress1 = (lc.LossCodesLevel == 1 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                ress2 = (lc.LossCodesLevel == 2 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                            }).ToList();

                            if (lossData.Count > 0)
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                foreach (var res in lossData)
                                {
                                    temp = temp.Replace("{{slno}}", slno);
                                    temp = temp.Replace("{{machinename}}", MachineName);
                                    temp = temp.Replace("{{mode}}", check1.Mode);
                                    temp = temp.Replace("{{rl1}}", res.ress1);
                                    temp = temp.Replace("{{rl2}}", res.ress2);
                                    temp = temp.Replace("{{rl3}}", res.ress3);
                                    temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                    temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                    strReplace = strReplace + temp;
                                }
                            }
                            else
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                temp = temp.Replace("{{slno}}", slno);
                                temp = temp.Replace("{{machinename}}", MachineName);
                                temp = temp.Replace("{{mode}}", check1.Mode);
                                temp = temp.Replace("{{rl1}}", "-");
                                temp = temp.Replace("{{rl2}}", "-");
                                temp = temp.Replace("{{rl3}}", "-");
                                temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                strReplace = strReplace + temp;
                            }
                            sl++;
                        }

                        htmlStr = htmlStr.Replace("{{strReplace}}", strReplace);
                        htmlStr = htmlStr.Replace("{{LossLevel}}", "2");

                        //string acceptUrl = configuration.GetSection("AppSettings").GetSection("AcceptURLDASDataCorr").Value;
                        //string rejectUrl = configuration.GetSection("AppSettings").GetSection("RejectURLDASDataCorr").Value;

                        //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLDASDataCorr").Value;
                        //string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLDASDataCorr").Value;

                        //string rejectSrc = rejectUrl + "CorrectedDate=" + correctedDate + "&machineId=" + machineId1 + "&tempId=" + tempModeIds + "&plantID=" + getMacDetails.PlantId + "&cellId=" + getMacDetails.CellId + "&shopId=" + getMacDetails.ShopId + "";
                        //string acceptSrc = acceptUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + tempModeIds + "&plantID=" + getMachineDet.PlantId + "&cellId=" + getMachineDet.CellId + "&shopId=" + getMachineDet.ShopId + "";

                        string macName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();

                        string TomailIds = " ";
                        string CcmailIds = " ";
                        string SecondLevelTomailIds = " ";
                        string SecondLevelCcmailIds = " ";

                        var mailDetails = (from wf in db.TblTcfApprovedMaster
                                           where wf.TcfModuleId == 3 && (wf.CellId == getMacDetails.CellId || wf.ShopId == getMacDetails.ShopId)
                                           select new
                                           {
                                               FirstApproverToList = wf.FirstApproverToList,
                                               FirstApproverCcList = wf.FirstApproverCcList,
                                               SecondApproverToList = wf.SecondApproverToList,
                                               SecondApproverCcList = wf.SecondApproverCcList
                                           }).ToList();

                        if (mailDetails.Count > 0)
                        {
                            foreach (var items in mailDetails)
                            {
                                TomailIds = items.FirstApproverToList;
                                CcmailIds = items.FirstApproverCcList;
                                SecondLevelTomailIds = items.SecondApproverToList;
                                SecondLevelCcmailIds = items.SecondApproverCcList;
                            }
                        }

                        string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                        "</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Approved For First Level</center></p></br><p>Thank you" +
                        "</p></br><p>Sincerely,</p><p>Dear</p></br></body></html>";

                        //bool ret = SendMail(message, SecondLevelTomailIds, SecondLevelCcmailIds, 1, macName, correctedDate);
                        bool ret = SendMail(htmlStr, TomailIds, CcmailIds, 0, macName, correctedDate, message);
                        //ret = SendMail(htmlStr, TomailIds, CcmailIds, 1, macName, correctedDate, null);

                        if (ret)
                        {
                            obj.isStatus = true;
                            obj.response = "Mail Sent Successfully";
                        }
                        #endregion
                    }
                }
                if (checkData3.Count > 0)
                {
                    if (mailInfo1.Count > 0)
                    {
                        #region mimics updation

                        UpdateMimicsDetails(machineID, correctedDate);

                        #endregion

                        UpdateInReport(machineID, correctedDate);

                        #region Sending Mail data to the second Level

                        int sl = 1;
                        string MachineName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();
                        string PlantName = db.Tblplant.Where(m => m.PlantId == getMacDetails.PlantId).Select(m => m.PlantName).FirstOrDefault();
                        string ShopName = db.Tblshop.Where(m => m.ShopId == getMacDetails.ShopId).Select(m => m.ShopName).FirstOrDefault();
                        string cellName = db.Tblcell.Where(m => m.CellId == getMacDetails.CellId).Select(m => m.CellName).FirstOrDefault();
                        int machineId1 = machineID;

                        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\DataLossTcfTemplate1.html");
                        string htmlStr = File.ReadAllText(reader);


                        //string[] ids1 = tempModeIds.Split(',');

                        string mainTemp = "<tr><td><span>{{slno}}</span> </td> <td><span>{{machinename}}</span> </td><td><span>{{mode}}" +
                            "</span> </td><td><span>{{rl1}}</span></td><td><span>{{rl2}}</span></td><td><span>{{rl3}}</span></  td><td><span>{{startTime}}</span></td><td><span>{{endTime}}</span></td></tr>";
                        string strReplace = "";

                        foreach (var item1 in checkData3)
                        {
                            var check1 = db.TblTempMode.Where(m => m.TempModeId == item1.TempModeId && m.IsApproved == 2 && m.ApproveLevel == 2).FirstOrDefault();

                            var lossData = (from wf in db.TblTempLiveLossOfEntry
                                            join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                            where wf.TempModeId == item1.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == machineID && wf.CorrectedDate == correctedDate
                                            select new
                                            {
                                                LossCodeId = lc.LossCodeId,
                                                LossCode = lc.LossCode,
                                                LossCodesLevel = lc.LossCodesLevel,
                                                LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                StartDateTime = wf.StartDateTime,
                                                EndDateTime = wf.EndDateTime,
                                                MachineId = wf.MachineId,
                                                MessageCodeId = wf.MessageCodeId,
                                                CorrectedDate = wf.CorrectedDate,
                                                TempLossId = wf.TempLossId,

                                                ress1 = (lc.LossCodesLevel == 1 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                ress2 = (lc.LossCodesLevel == 2 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                            }).ToList();

                            if (lossData.Count > 0)
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                foreach (var res in lossData)
                                {
                                    temp = temp.Replace("{{slno}}", slno);
                                    temp = temp.Replace("{{machinename}}", MachineName);
                                    temp = temp.Replace("{{mode}}", check1.Mode);
                                    temp = temp.Replace("{{rl1}}", res.ress1);
                                    temp = temp.Replace("{{rl2}}", res.ress2);
                                    temp = temp.Replace("{{rl3}}", res.ress3);
                                    temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                    temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                    strReplace = strReplace + temp;
                                }
                            }
                            else
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                temp = temp.Replace("{{slno}}", slno);
                                temp = temp.Replace("{{machinename}}", MachineName);
                                temp = temp.Replace("{{mode}}", check1.Mode);
                                temp = temp.Replace("{{rl1}}", "-");
                                temp = temp.Replace("{{rl2}}", "-");
                                temp = temp.Replace("{{rl3}}", "-");
                                temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                strReplace = strReplace + temp;
                            }
                            sl++;
                        }

                        htmlStr = htmlStr.Replace("{{strReplace}}", strReplace);
                        htmlStr = htmlStr.Replace("{{LossLevel}}", "2");

                        //string acceptUrl = configuration.GetSection("AppSettings").GetSection("AcceptURLDASDataCorr").Value;
                        //string rejectUrl = configuration.GetSection("AppSettings").GetSection("RejectURLDASDataCorr").Value;

                        //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLDASDataCorr").Value;
                        //string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLDASDataCorr").Value;

                        //string rejectSrc = rejectUrl + "CorrectedDate=" + correctedDate + "&machineId=" + machineId1 + "&tempId=" + tempModeIds + "&plantID=" + getMacDetails.PlantId + "&cellId=" + getMacDetails.CellId + "&shopId=" + getMacDetails.ShopId + "";
                        //string acceptSrc = acceptUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + tempModeIds + "&plantID=" + getMachineDet.PlantId + "&cellId=" + getMachineDet.CellId + "&shopId=" + getMachineDet.ShopId + "";

                        string macName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();

                        string TomailIds = " ";
                        string CcmailIds = " ";
                        string SecondLevelTomailIds = " ";
                        string SecondLevelCcmailIds = " ";

                        var mailDetails = (from wf in db.TblTcfApprovedMaster
                                           where wf.TcfModuleId == 3 && (wf.CellId == getMacDetails.CellId || wf.ShopId == getMacDetails.ShopId)
                                           select new
                                           {
                                               FirstApproverToList = wf.FirstApproverToList,
                                               FirstApproverCcList = wf.FirstApproverCcList,
                                               SecondApproverToList = wf.SecondApproverToList,
                                               SecondApproverCcList = wf.SecondApproverCcList
                                           }).ToList();

                        if (mailDetails.Count > 0)
                        {
                            foreach (var items in mailDetails)
                            {
                                TomailIds = items.FirstApproverToList;
                                CcmailIds = items.FirstApproverCcList;
                                SecondLevelTomailIds = items.SecondApproverToList;
                                SecondLevelCcmailIds = items.SecondApproverCcList;
                            }
                        }

                        string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                        "</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Approved For Second Level</center></p></br><p>Thank you" +
                        "</p></br><p>Sincerely,</p><p>Dear</p></br></body></html>";

                        //bool ret = SendMail(message, SecondLevelTomailIds, SecondLevelCcmailIds, 1, macName, correctedDate);
                        bool ret = SendMail(htmlStr, SecondLevelTomailIds, SecondLevelCcmailIds, 0, macName, correctedDate, message);
                        ret = SendMail(message, TomailIds, CcmailIds, 1, macName, correctedDate, null);

                        if (ret)
                        {
                            obj.isStatus = true;
                            obj.response = "Mail Sent Successfully";
                        }

                        #endregion
                    }
                }

                #endregion

                #region After 1st Approved Send Mail

                //var checkData = db.TblTempMode.Where(m => m.MachineId == machineID && m.ApproveLevel == 1 && m.CorrectedDate == correctedDate).FirstOrDefault();
                //if (checkData != null)
                //{
                //    var getMachineDetails1 = (from wf in db.Tblmachinedetails
                //                              where wf.MachineId == machineID && wf.IsDeleted == 0
                //                              select new
                //                              {
                //                                  PlantId = wf.PlantId,
                //                                  ShopId = wf.ShopId,
                //                                  CellId = wf.CellId
                //                              }).FirstOrDefault();

                //    string machineName1 = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();

                //    string TomailIds = " ";
                //    string CcmailIds = " ";
                //    var mailDetails = (from wf in db.TblTcfApprovedMaster
                //                       where wf.TcfModuleId == 3 && (wf.CellId == getMachineDetails1.CellId || wf.ShopId == getMachineDetails1.ShopId)
                //                       select new
                //                       {
                //                           FirstApproverToList = wf.FirstApproverToList,
                //                           FirstApproverCcList = wf.FirstApproverCcList
                //                       }).ToList();

                //    if (mailDetails.Count > 0)
                //    {
                //        foreach (var items in mailDetails)
                //        {
                //            TomailIds = items.FirstApproverToList;
                //            CcmailIds = items.FirstApproverCcList;
                //        }
                //    }

                //    string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                //    "</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Approved For First Level</center></p></br><p>Thank you" +
                //    "</p></br><p>Sincerely,</p><p>Dear</p></br></body></html>";

                //    bool ret = SendMail(message, TomailIds, CcmailIds, 1, machineName1, correctedDate);
                //    if (ret)
                //    {
                //        obj.isStatus = true;
                //        obj.response = "Mail Sent Successfully";
                //    }
                //}
                #endregion

                #region After 1st Approved Send Mail

                //var checkData1 = db.TblTempMode.Where(m => m.MachineId == machineID && m.ApproveLevel == 2 && m.CorrectedDate == correctedDate).ToList();
                //if (checkData1.Count > 0)
                //{
                //    var getMachineDetails = (from wf in db.Tblmachinedetails
                //                             where wf.MachineId == machineID && wf.IsDeleted == 0
                //                             select new
                //                             {
                //                                 PlantId = wf.PlantId,
                //                                 ShopId = wf.ShopId,
                //                                 CellId = wf.CellId
                //                             }).FirstOrDefault();


                //    int sl = 1;
                //    string MachineName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();
                //    string PlantName = db.Tblplant.Where(m => m.PlantId == getMachineDetails.PlantId).Select(m => m.PlantName).FirstOrDefault();
                //    string ShopName = db.Tblshop.Where(m => m.ShopId == getMachineDetails.ShopId).Select(m => m.ShopName).FirstOrDefault();
                //    string cellName = db.Tblcell.Where(m => m.CellId == getMachineDetails.CellId).Select(m => m.CellName).FirstOrDefault();
                //    int machineId1 = machineID;

                //    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\DataLossTcfTemplate1.html");
                //    string htmlStr = File.ReadAllText(reader);


                //    //string[] ids1 = tempModeIds.Split(',');

                //    string mainTemp = "<tr><td><span>{{slno}}</span> </td> <td><span>{{machinename}}</span> </td><td><span>{{mode}}" +
                //        "</span> </td><td><span>{{rl1}}</span></td><td><span>{{rl2}}</span></td><td><span>{{rl3}}</span></  td><td><span>{{startTime}}</span></td><td><span>{{endTime}}</span></td></tr>";
                //    string strReplace = "";

                //    foreach (var item1 in checkData1)
                //    {
                //        var check1 = db.TblTempMode.Where(m => m.TempModeId == item1.TempModeId && m.IsApproved == 2 && m.ApproveLevel == 2).FirstOrDefault();

                //        var lossData = (from wf in db.TblTempLiveLossOfEntry
                //                        join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                //                        where wf.TempModeId == item1.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == machineID && wf.CorrectedDate == correctedDate
                //                        select new
                //                        {
                //                            LossCodeId = lc.LossCodeId,
                //                            LossCode = lc.LossCode,
                //                            LossCodesLevel = lc.LossCodesLevel,
                //                            LossCodesLevel1Id = lc.LossCodesLevel1Id,
                //                            LossCodesLevel2Id = lc.LossCodesLevel2Id,
                //                            StartDateTime = wf.StartDateTime,
                //                            EndDateTime = wf.EndDateTime,
                //                            MachineId = wf.MachineId,
                //                            MessageCodeId = wf.MessageCodeId,
                //                            CorrectedDate = wf.CorrectedDate,
                //                            TempLossId = wf.TempLossId,

                //                            ress1 = (lc.LossCodesLevel == 1 ?
                //                             lc.LossCode :
                //                             db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                //                            ress2 = (lc.LossCodesLevel == 2 ?
                //                             lc.LossCode :
                //                             db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                //                            ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                //                        }).ToList();

                //        if (lossData.Count > 0)
                //        {
                //            string temp = mainTemp;
                //            string slno = Convert.ToString(sl);
                //            foreach (var res in lossData)
                //            {
                //                temp = temp.Replace("{{slno}}", slno);
                //                temp = temp.Replace("{{machinename}}", MachineName);
                //                temp = temp.Replace("{{mode}}", check1.Mode);
                //                temp = temp.Replace("{{rl1}}", res.ress1);
                //                temp = temp.Replace("{{rl2}}", res.ress2);
                //                temp = temp.Replace("{{rl3}}", res.ress3);
                //                temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                //                temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                //                strReplace = strReplace + temp;
                //            }
                //        }
                //        else
                //        {
                //            string temp = mainTemp;
                //            string slno = Convert.ToString(sl);
                //            temp = temp.Replace("{{slno}}", slno);
                //            temp = temp.Replace("{{machinename}}", MachineName);
                //            temp = temp.Replace("{{mode}}", check1.Mode);
                //            temp = temp.Replace("{{rl1}}", "-");
                //            temp = temp.Replace("{{rl2}}", "-");
                //            temp = temp.Replace("{{rl3}}", "-");
                //            temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                //            temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                //            strReplace = strReplace + temp;
                //        }
                //        sl++;
                //    }

                //    htmlStr = htmlStr.Replace("{{strReplace}}", strReplace);
                //    htmlStr = htmlStr.Replace("{{LossLevel}}", "2");

                //    //string acceptUrl = configuration.GetSection("AppSettings").GetSection("AcceptURLDASDataCorr").Value;
                //    //string rejectUrl = configuration.GetSection("AppSettings").GetSection("RejectURLDASDataCorr").Value;

                //    //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLDASDataCorr").Value;
                //    string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLDASDataCorr").Value;

                //    string rejectSrc = rejectUrl + "CorrectedDate=" + correctedDate + "&machineId=" + machineId1 + "&tempId=" + tempModeIds + "&plantID=" + getMachineDetails.PlantId + "&cellId=" + getMachineDetails.CellId + "&shopId=" + getMachineDetails.ShopId + "";
                //    //string acceptSrc = acceptUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + tempModeIds + "&plantID=" + getMachineDet.PlantId + "&cellId=" + getMachineDet.CellId + "&shopId=" + getMachineDet.ShopId + "";

                //    string macName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();

                //    string TomailIds = " ";
                //    string CcmailIds = " ";
                //    string SecondLevelTomailIds = " ";
                //    string SecondLevelCcmailIds = " ";

                //    var mailDetails = (from wf in db.TblTcfApprovedMaster
                //                       where wf.TcfModuleId == 3 && (wf.CellId == getMachineDetails.CellId || wf.ShopId == getMachineDetails.ShopId)
                //                       select new
                //                       {
                //                           FirstApproverToList = wf.FirstApproverToList,
                //                           FirstApproverCcList = wf.FirstApproverCcList,
                //                           SecondApproverToList = wf.SecondApproverToList,
                //                           SecondApproverCcList = wf.SecondApproverCcList
                //                       }).ToList();

                //    if (mailDetails.Count > 0)
                //    {
                //        foreach (var items in mailDetails)
                //        {
                //            TomailIds = items.FirstApproverToList;
                //            CcmailIds = items.FirstApproverCcList;
                //            SecondLevelTomailIds = items.SecondApproverToList;
                //            SecondLevelCcmailIds = items.SecondApproverCcList;
                //        }
                //    }

                //    string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                //    "</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Approved For Second Level</center></p></br><p>Thank you" +
                //    "</p></br><p>Sincerely,</p><p>Dear</p></br></body></html>";

                //    //bool ret = SendMail(message, SecondLevelTomailIds, SecondLevelCcmailIds, 1, macName, correctedDate);
                //    bool ret = SendMail(htmlStr, SecondLevelTomailIds, SecondLevelCcmailIds, 1, macName, correctedDate, message);
                //    ret = SendMail(message, TomailIds, CcmailIds, 1, macName, correctedDate, null);

                //    if (ret)
                //    {
                //        obj.isStatus = true;
                //        obj.response = "Mail Sent Successfully";
                //    }
                //}
                #endregion

                #region sending mail if second approver list is present

                var getMachineDet = (from wf in db.Tblmachinedetails
                                     where wf.MachineId == machineID && wf.IsDeleted == 0
                                     select new
                                     {
                                         PlantId = wf.PlantId,
                                         ShopId = wf.ShopId,
                                         CellId = wf.CellId
                                     }).FirstOrDefault();

                string machineName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();

                string[] ids2 = tempModeIds.Split(',');
                int machineId = 0;
                string CorrectedDate = " ";
                foreach (var item in ids2)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsApproved == 2 && m.ApproveLevel == 1).FirstOrDefault();
                    if (check != null)
                    {
                        machineId = check.MachineId;
                        CorrectedDate = check.CorrectedDate;
                    }
                }

                var getMachineDet1 = (from wf in db.Tblmachinedetails
                                      where wf.MachineId == machineId && wf.IsDeleted == 0
                                      select new
                                      {
                                          PlantId = wf.PlantId,
                                          ShopId = wf.ShopId,
                                          CellId = wf.CellId
                                      }).FirstOrDefault();

                if (getMachineDet1 != null)
                {
                    var mailDetails1 = (from wf in db.TblTcfApprovedMaster
                                        where wf.TcfModuleId == 3 && (wf.CellId == getMachineDet1.CellId || wf.ShopId == getMachineDet1.ShopId) && wf.SecondApproverToList != null && wf.SecondApproverToList != ""
                                        && wf.SecondApproverCcList != null && wf.SecondApproverCcList != ""
                                        select new
                                        {
                                            SecondApproverToList = wf.SecondApproverToList,
                                            SecondApproverCcList = wf.SecondApproverCcList
                                        }).ToList();

                    string SecondApproverToListIds = " ";
                    string SecondApproverCcListIds = " ";

                    if (mailDetails1.Count > 0)
                    {

                        //SpliDurationDAL spliDurationDAL = new SpliDurationDAL(db, configuration);
                        //spliDurationDAL.SendMailToApproveOrReject(sendMailDetails);

                        int sl = 1;
                        string MachineName = db.Tblmachinedetails.Where(m => m.MachineId == machineID).Select(m => m.MachineInvNo).FirstOrDefault();
                        string PlantName = db.Tblplant.Where(m => m.PlantId == getMachineDet1.PlantId).Select(m => m.PlantName).FirstOrDefault();
                        string ShopName = db.Tblshop.Where(m => m.ShopId == getMachineDet1.ShopId).Select(m => m.ShopName).FirstOrDefault();
                        string cellName = db.Tblcell.Where(m => m.CellId == getMachineDet1.CellId).Select(m => m.CellName).FirstOrDefault();
                        int machineId1 = machineID;

                        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\DataLossTcfTemplate1.html");
                        string htmlStr = File.ReadAllText(reader);


                        string[] ids1 = tempModeIds.Split(',');

                        string mainTemp = "<tr><td><span>{{slno}}</span> </td> <td><span>{{machinename}}</span> </td><td><span>{{mode}}" +
                            "</span> </td><td><span>{{rl1}}</span></td><td><span>{{rl2}}</span></td><td><span>{{rl3}}</span></  td><td><span>{{startTime}}</span></td><td><span>{{endTime}}</span></td></tr>";
                        string strReplace = "";

                        foreach (var item1 in ids1)
                        {
                            int tempModeID = Convert.ToInt32(item1);
                            var check1 = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 2 && m.ApproveLevel == 1).FirstOrDefault();
                            if (check1 != null)
                            {
                                check1.IsApproved1 = 1;
                                db.SaveChanges();
                            }

                            var lossData = (from wf in db.TblTempLiveLossOfEntry
                                            join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                            where wf.TempModeId == tempModeID && wf.DoneWithRow == 1 && wf.MachineId == machineID && wf.CorrectedDate == CorrectedDate
                                            select new
                                            {
                                                LossCodeId = lc.LossCodeId,
                                                LossCode = lc.LossCode,
                                                LossCodesLevel = lc.LossCodesLevel,
                                                LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                StartDateTime = wf.StartDateTime,
                                                EndDateTime = wf.EndDateTime,
                                                MachineId = wf.MachineId,
                                                MessageCodeId = wf.MessageCodeId,
                                                CorrectedDate = wf.CorrectedDate,
                                                TempLossId = wf.TempLossId,

                                                ress1 = (lc.LossCodesLevel == 1 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                ress2 = (lc.LossCodesLevel == 2 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                            }).ToList();

                            if (lossData.Count > 0)
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                foreach (var res in lossData)
                                {
                                    temp = temp.Replace("{{slno}}", slno);
                                    temp = temp.Replace("{{machinename}}", MachineName);
                                    temp = temp.Replace("{{mode}}", check1.Mode);
                                    temp = temp.Replace("{{rl1}}", res.ress1);
                                    temp = temp.Replace("{{rl2}}", res.ress2);
                                    temp = temp.Replace("{{rl3}}", res.ress3);
                                    temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                    temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                    strReplace = strReplace + temp;
                                }
                            }
                            else
                            {
                                string temp = mainTemp;
                                string slno = Convert.ToString(sl);
                                temp = temp.Replace("{{slno}}", slno);
                                temp = temp.Replace("{{machinename}}", MachineName);
                                temp = temp.Replace("{{mode}}", check1.Mode);
                                temp = temp.Replace("{{rl1}}", "-");
                                temp = temp.Replace("{{rl2}}", "-");
                                temp = temp.Replace("{{rl3}}", "-");
                                temp = temp.Replace("{{startTime}}", check1.StartTime.ToString());
                                temp = temp.Replace("{{endTime}}", check1.EndTime.ToString());
                                strReplace = strReplace + temp;
                            }
                            sl++;
                        }

                        htmlStr = htmlStr.Replace("{{strReplace}}", strReplace);
                        htmlStr = htmlStr.Replace("{{LossLevel}}", "2");

                        //string acceptUrl = configuration.GetSection("AppSettings").GetSection("AcceptURLDASDataCorr").Value;
                        //string rejectUrl = configuration.GetSection("AppSettings").GetSection("RejectURLDASDataCorr").Value;

                        //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLDASDataCorr").Value;
                        string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLDASDataCorr").Value;

                        string rejectSrc = rejectUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId1 + "&tempId=" + tempModeIds + "&plantID=" + getMachineDet.PlantId + "&cellId=" + getMachineDet.CellId + "&shopId=" + getMachineDet.ShopId + "&checked="+ tempModeIds + "";
                        //string acceptSrc = acceptUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + tempModeIds + "&plantID=" + getMachineDet.PlantId + "&cellId=" + getMachineDet.CellId + "&shopId=" + getMachineDet.ShopId + "";

                        string SecondApproverToMailIds = " ";
                        string SecondApproverCcMailIds = " ";

                        foreach (var items in mailDetails1)
                        {
                            if (items.SecondApproverToList != null && items.SecondApproverCcList != null)
                            {
                                SecondApproverToMailIds = items.SecondApproverToList;
                                SecondApproverCcMailIds = items.SecondApproverCcList;
                            }
                        }

                        //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                        htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                        bool ret = SendMail(htmlStr, SecondApproverToMailIds, SecondApproverCcMailIds, 1, MachineName, CorrectedDate, null);

                        //message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link   rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                        //"</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Approved For Second Approval</center></p></ br><p>Thank you" +"</p></br><p>Sincerely,</p><p>Dear</p></br></body></html>";

                        //ret = SendMail(message, SecondApproverToMailIds, SecondApproverCcMailIds, 1, machineName, correctedDate);

                        if (ret)
                        {
                            obj.isStatus = true;
                            obj.response = "Second Approval Mail Sent Successfully";
                        }
                    }
                }
                else
                {
                    obj.isStatus = true;
                }
                #endregion
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        ///// <summary>
        ///// Update Mode Details
        ///// </summary>
        ///// <param name="ModeId"></param>
        ///// <param name="EndTime"></param>
        ///// <param name="Mode"></param>
        ///// <returns></returns>
        //public CommonResponse1 UpdateModeDetails(List<ModeStartEndDateTime> data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        int i = 0;
        //        List<ModeStartEndDateTime> modeStartEndDateTimeList = new List<ModeStartEndDateTime>();
        //        foreach (var item in data)
        //        {
        //            i++;
        //            modeStartEndDateTimeList.Add(item);
        //            if (data.Count == i)
        //            {
        //                var check = db.Tblmode.Where(m => m.ModeId == item.modeId).FirstOrDefault();

        //                if (check != null)
        //                {

        //                    DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
        //                    DateTime StartDateTime = Convert.ToDateTime(check.StartTime);

        //                    string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    string[] ids = dt1.Split();
        //                    string endDate = ids[0];
        //                    string endTime = ids[1];
        //                    string endTimeLast = ids[0] + " " + item.endTime;
        //                    DateTime endT = Convert.ToDateTime(endTimeLast);
        //                    if (endT < EndDateTime && endT > StartDateTime)
        //                    {
        //                        #region Update old row
        //                        check.EndTime = endT;
        //                        check.Mode = "PowerOff";
        //                        var durationInSec = (endT - StartDateTime).TotalSeconds;
        //                        check.DurationInSec = Convert.ToInt32(durationInSec);
        //                        db.SaveChanges();
        //                        #endregion

        //                        #region update in the Live mode more than less than 5 days
        //                        var dbCheck = db.Tbllivemodedb.Where(m => m.ModeId == item.modeId).FirstOrDefault();
        //                        if (dbCheck != null)
        //                        {
        //                            DateTime prevCorrectedDate = Convert.ToDateTime(dbCheck.CorrectedDate);
        //                            DateTime todayDate = DateTime.Now;
        //                            int diffDate = Convert.ToInt32((todayDate - prevCorrectedDate).TotalDays);
        //                            if (diffDate < 5)
        //                            {
        //                                DateTime StartDateTime1 = Convert.ToDateTime(dbCheck.StartTime);
        //                                dbCheck.EndTime = endT;
        //                                dbCheck.Mode = "PowerOff";
        //                                var durationInSec1 = (endT - StartDateTime1).TotalSeconds;
        //                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec1);
        //                                db.SaveChanges();
        //                            }
        //                        }

        //                        #endregion

        //                        #region add new row in live mode table which we are going to insert in mode table
        //                        Tbllivemodedb tbllivemodedbNew = new Tbllivemodedb();
        //                        tbllivemodedbNew.MachineId = check.MachineId;
        //                        tbllivemodedbNew.Mode = "PowerOff";
        //                        tbllivemodedbNew.ColorCode = "blue";
        //                        tbllivemodedbNew.InsertedOn = DateTime.Now;
        //                        tbllivemodedbNew.InsertedBy = 1;
        //                        tbllivemodedbNew.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
        //                        tbllivemodedbNew.IsDeleted = 0;
        //                        tbllivemodedbNew.StartTime = endT;
        //                        tbllivemodedbNew.EndTime = EndDateTime;
        //                        tbllivemodedbNew.IsCompleted = 1;
        //                        var durationInSec11 = (EndDateTime - endT).TotalSeconds;
        //                        tbllivemodedbNew.DurationInSec = Convert.ToInt32(durationInSec11);
        //                        db.Tbllivemodedb.Add(tbllivemodedbNew);
        //                        db.SaveChanges();
        //                        #endregion

        //                        #region add new row in mode table which we have inserted in live mode table in last region
        //                        Tblmode tblmode = new Tblmode();
        //                        tblmode.MachineId = check.MachineId;
        //                        tblmode.ColorCode = tbllivemodedbNew.ColorCode;
        //                        tblmode.ModeId = tbllivemodedbNew.ModeId;
        //                        tblmode.Mode = tbllivemodedbNew.Mode;
        //                        tblmode.IsDeleted = 0;
        //                        tblmode.InsertedOn = DateTime.Now;
        //                        tblmode.InsertedBy = 1;
        //                        tblmode.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
        //                        tblmode.StartTime = endT;
        //                        tblmode.EndTime = EndDateTime;
        //                        tblmode.IsCompleted = 1;
        //                        var durationInSec2 = (EndDateTime - endT).TotalSeconds;
        //                        tblmode.DurationInSec = Convert.ToInt32(durationInSec2);
        //                        db.Tblmode.Add(tblmode);
        //                        db.SaveChanges();
        //                        #endregion

        //                        #region Response assiging
        //                        ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
        //                        modeStartEndDateTime.startDate = endDate;
        //                        modeStartEndDateTime.startTime = item.endTime;
        //                        modeStartEndDateTime.endDate = endDate;
        //                        modeStartEndDateTime.endTime = endTime;
        //                        modeStartEndDateTime.modeId = tblmode.ModeId;
        //                        modeStartEndDateTimeList.Add(modeStartEndDateTime);
        //                        obj.isStatus = true;
        //                        obj.response = modeStartEndDateTimeList;
        //                        #endregion

        //                    }
        //                    else
        //                    {
        //                        obj.isStatus = true;
        //                        obj.response = "Please Enter the EndDateTime with in" + EndDateTime;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        ///// <summary>
        ///// Add Reasons for mode Details
        ///// </summary>
        ///// <param name="data"></param>
        ///// <returns></returns>
        //public CommonResponse1 AddLossDetails(List<AddReasonsDetails> data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {

        //        foreach (var item in data)
        //        {
        //            DateTime Time = DateTime.Now;
        //            TimeSpan Tm = new TimeSpan(Time.Hour, Time.Minute, Time.Second);
        //            var ShiftDetails = db.TblshiftMstr.Where(m => m.StartTime <= Tm && m.EndTime >= Tm).FirstOrDefault();
        //            string Shift = "C";
        //            if (ShiftDetails != null)
        //            {
        //                Shift = ShiftDetails.ShiftName;
        //            }

        //            if (item.lossCodeId != 0)
        //            {
        //                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == item.lossCodeId).FirstOrDefault();

        //                var modeData = db.TblTempMode.Where(m => m.TempModeId == item.tempModeId && m.IsApproved == 0).FirstOrDefault();
        //                if (modeData != null)
        //                {
        //                    if (item.mode == "Idle")
        //                    {
        //                        modeData.ColorCode = "yellow";
        //                        modeData.Mode = "IDLE";
        //                    }
        //                    if (item.mode == "Breakdown")
        //                    {
        //                        modeData.ColorCode = "red";
        //                        modeData.Mode = "BREAKDOWN";
        //                    }
        //                    if (item.mode == "Running")
        //                    {
        //                        modeData.Mode = "Running";
        //                        modeData.ColorCode = "green";
        //                    }
        //                    if (item.mode == "PowerOff")
        //                    {
        //                        modeData.Mode = "PowerOff";
        //                        modeData.ColorCode = "blue";
        //                    }

        //                    modeData.Mode = modeData.Mode;
        //                    modeData.ColorCode = modeData.ColorCode;
        //                    modeData.OverAllSaved = 1;
        //                    db.SaveChanges();
        //                }

        //                var check = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == modeData.TempModeId).FirstOrDefault();
        //                if (check == null)
        //                {
        //                    TblTempLiveLossOfEntry tblTempLiveLossOfEntry = new TblTempLiveLossOfEntry();
        //                    tblTempLiveLossOfEntry.StartDateTime = modeData.StartTime;
        //                    tblTempLiveLossOfEntry.EntryTime = modeData.StartTime;
        //                    tblTempLiveLossOfEntry.EndDateTime = modeData.EndTime;
        //                    tblTempLiveLossOfEntry.CorrectedDate = modeData.CorrectedDate;
        //                    tblTempLiveLossOfEntry.MachineId = item.machineId;
        //                    tblTempLiveLossOfEntry.Shift = Shift;
        //                    tblTempLiveLossOfEntry.MessageCodeId = item.lossCodeId;
        //                    tblTempLiveLossOfEntry.MessageCode = lossdata.LossCode;
        //                    tblTempLiveLossOfEntry.MessageDesc = lossdata.LossCodeDesc;
        //                    tblTempLiveLossOfEntry.IsUpdate = 1;
        //                    tblTempLiveLossOfEntry.DoneWithRow = 1;
        //                    tblTempLiveLossOfEntry.IsStart = 0;
        //                    tblTempLiveLossOfEntry.IsScreen = 0;
        //                    tblTempLiveLossOfEntry.ForRefresh = 0;
        //                    tblTempLiveLossOfEntry.TempModeId = modeData.TempModeId;
        //                    tblTempLiveLossOfEntry.ModeId = modeData.ModeId;
        //                    db.TblTempLiveLossOfEntry.Add(tblTempLiveLossOfEntry);
        //                    db.SaveChanges();
        //                    obj.isStatus = true;
        //                    obj.response = "Added Successfully";
        //                }
        //                else
        //                {
        //                    check.StartDateTime = modeData.StartTime;
        //                    check.EntryTime = modeData.StartTime;
        //                    check.EndDateTime = modeData.EndTime;
        //                    check.CorrectedDate = modeData.CorrectedDate;
        //                    check.MachineId = item.machineId;
        //                    check.Shift = Shift;
        //                    check.MessageCodeId = item.lossCodeId;
        //                    check.MessageCode = lossdata.LossCode;
        //                    check.MessageDesc = lossdata.LossCodeDesc;
        //                    check.IsUpdate = 1;
        //                    check.DoneWithRow = 1;
        //                    check.IsStart = 0;
        //                    check.IsScreen = 0;
        //                    check.ForRefresh = 0;
        //                    check.TempModeId = modeData.TempModeId;
        //                    check.ModeId = modeData.ModeId;
        //                    db.SaveChanges();
        //                    obj.isStatus = true;
        //                    obj.response = "Updated Successfully";
        //                }
        //            }
        //            else
        //            {
        //                var modeData = db.TblTempMode.Where(m => m.TempModeId == item.tempModeId && m.IsApproved == 0).FirstOrDefault();
        //                if (modeData != null)
        //                {
        //                    if (item.mode == "Idle")
        //                    {
        //                        modeData.ColorCode = "yellow";
        //                        modeData.Mode = "IDLE";
        //                    }
        //                    if (item.mode == "Breakdown")
        //                    {
        //                        modeData.ColorCode = "red";
        //                        modeData.Mode = "BREAKDOWN";
        //                    }
        //                    if (item.mode == "Running")
        //                    {
        //                        modeData.Mode = "Running";
        //                        modeData.ColorCode = "green";
        //                    }
        //                    if (item.mode == "PowerOff")
        //                    {
        //                        modeData.Mode = "PowerOff";
        //                        modeData.ColorCode = "blue";
        //                    }

        //                    modeData.Mode = modeData.Mode;
        //                    modeData.ColorCode = modeData.ColorCode;
        //                    modeData.OverAllSaved = 1;
        //                    db.SaveChanges();
        //                    obj.isStatus = true;
        //                    obj.response = "Updated Successfully";
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        /// <summary>
        /// Idle Code Resons for Level 1
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetIdleResonLevel1(string mode)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (mode == "Idle")
                {

                    var level1Data = (from wf in db.Tbllossescodes
                                      where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType != "NoCode" && wf.MessageType != "BREAKDOWN" && wf.MessageType != "PM"
                                      select new
                                      {
                                          LossCodeId = wf.LossCodeId,
                                          LossCode = wf.LossCode
                                      }).ToList();

                    if (level1Data.Count != 0)
                    {
                        obj.isStatus = true;
                        obj.response = level1Data;
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "No Items Found";
                    }
                }
                else if (mode == "Breakdown")
                {
                    var level1Data = (from wf in db.Tbllossescodes
                                      where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType == "BREAKDOWN" && wf.LossCode != "9999"
                                      select new
                                      {
                                          LossCode = wf.LossCode,
                                          LossCodeId = wf.LossCodeId,
                                      }).ToList();
                    if (level1Data.Count > 0)
                    {
                        obj.isStatus = true;
                        obj.response = level1Data;
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "No Items Found";
                    }
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Idle Code Resons for Level 2
        /// </summary>
        /// <param name="LossCodeID"></param>
        /// <returns></returns>
        public CommonResponse1 GetIdleResonLevel2(int LossCodeID)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var level2Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == LossCodeID && wf.LossCodesLevel == 2 && wf.LossCodesLevel2Id == null && wf.MessageType != "BREAKDOWN"
                                  select new
                                  {
                                      LossCodeId = wf.LossCodeId,
                                      LossCode = wf.LossCode,
                                  }).ToList();

                var level3Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == LossCodeID && wf.LossCodesLevel2Id == null && wf.MessageType == "BREAKDOWN"
                                  select new
                                  {
                                      LossCode = wf.LossCode,
                                      LossCodeId = wf.LossCodeId,
                                  }).ToList();

                if (level2Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level2Data;
                }
                else if (level3Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level3Data;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Idle Code Resons for Level 3
        /// </summary>
        /// <param name="LossCodeID"></param>
        /// <returns></returns>
        public CommonResponse1 GetIdleResonLevel3(int LossCodeID)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var level3Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == LossCodeID && wf.LossCodesLevel == 3 && wf.MessageType != "BREAKDOWN"
                                  select new
                                  {
                                      LossCodeId = wf.LossCodeId,
                                      LossCode = wf.LossCode
                                  }).ToList();

                var level2Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == LossCodeID && wf.MessageType == "BREAKDOWN"
                                  select new
                                  {
                                      LossCode = wf.LossCode,
                                      LossCodeId = wf.LossCodeId

                                  }).ToList();

                if (level3Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level3Data;
                }
                else if (level2Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level2Data;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
                return obj;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        ///// <summary>
        ///// BreakDown Code Resons for Level 1
        ///// </summary>
        ///// <returns></returns>
        //public CommonResponse1 GetBreakDownReasonLevel1()
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        var level1Data = (from wf in db.Tbllossescodes
        //                          where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType == "BREAKDOWN" && wf.LossCode != "9999"
        //                          select new
        //                          {
        //                              LossCode = wf.LossCode,
        //                              LossCodeId = wf.LossCodeId,
        //                          }).ToList();
        //        if (level1Data.Count > 0)
        //        {
        //            obj.isStatus = true;
        //            obj.response = level1Data;
        //        }
        //        else
        //        {
        //            obj.isStatus = false;
        //            obj.response = "No Items Found";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        ///// <summary>
        ///// BreakDown Code Resons for Level 2
        ///// </summary>
        ///// <param name="LossCodeID"></param>
        ///// <returns></returns>
        //public CommonResponse1 GetBreakDownReasonLevel2(int LossCodeID)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        var level2Data = (from wf in db.Tbllossescodes
        //                          where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == LossCodeID && wf.LossCodesLevel2Id == null && wf.MessageType == "BREAKDOWN"
        //                          select new
        //                          {
        //                              LossCode = wf.LossCode,
        //                              LossCodeId = wf.LossCodeId,
        //                          }).ToList();
        //        if (level2Data.Count == 0)
        //        {
        //            obj.isStatus = false;
        //            obj.response = "No Items Found";
        //        }
        //        else
        //        {
        //            obj.isStatus = true;
        //            obj.response = level2Data;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        /// <summary>
        /// BreakDown Code Resons for Level 3
        /// </summary>
        /// <param name="LossCodeID"></param>
        /// <returns></returns>
        //public CommonResponse1 GetBreakDownReasonLevel3(int LossCodeID)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        var level3Data = (from wf in db.Tbllossescodes
        //                          where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == LossCodeID && wf.MessageType == "BREAKDOWN"
        //                          select new
        //                          {
        //                              LossCode = wf.LossCode,
        //                              LossCodeId = wf.LossCodeId

        //                          }).ToList();

        //        if (level3Data.Count == 0)
        //        {
        //            obj.isStatus = false;
        //            obj.response = "No Items Found";
        //        }
        //        else
        //        {
        //            obj.isStatus = true;
        //            obj.response = level3Data;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        /// <summary>
        /// Cancel the duration 
        /// </summary>
        /// <param name="tempModeId"></param>
        /// <returns></returns>
        public GeneralResponse1 DeleteSplitDuration(int tempModeId, int modeId)
        {
            GeneralResponse1 obj = new GeneralResponse1();
            try
            {
                int prevTempModeId = 0;
                var check = db.TblTempMode.Where(m => m.ModeId == modeId && m.IsDeleted == 0).ToList();
                for (int i = 0; i < check.Count; i++)
                {
                    if (check[i].TempModeId.ToString().Contains(tempModeId.ToString()))
                    {
                        if ((i - 1) < 0)
                        {
                            prevTempModeId = check[i + 1].TempModeId;

                            var newTempMode = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();

                            var dbCheck = db.TblTempMode.Where(m => m.TempModeId == prevTempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                dbCheck.StartTime = newTempMode.StartTime;
                                dbCheck.EndTime = dbCheck.EndTime;
                                DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                var durationInSec = (EndDateTime - StartDateTime).TotalSeconds;
                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                                db.SaveChanges();

                                db.TblTempMode.Remove(newTempMode);
                                db.SaveChanges();

                            }

                            var tempData = db.TblTempMode.Where(m => m.ModeId == dbCheck.ModeId && m.MachineId == dbCheck.MachineId & m.IsSaved == 0).ToList();
                            List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();
                            foreach (var item1 in tempData)
                            {
                                DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
                                DateTime StartDateTime = Convert.ToDateTime(item1.StartTime);

                                string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids = dt.Split();
                                string endDate = ids[0];
                                string endTime = ids[1];

                                string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids1 = dt1.Split();
                                string startDate = ids1[0];
                                string startTime = ids1[1];

                                ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                                modeStartEndDateTime.startDate = startDate;
                                modeStartEndDateTime.startTime = startTime;
                                modeStartEndDateTime.endDate = endDate;
                                modeStartEndDateTime.endTime = endTime;
                                modeStartEndDateTime.modeId = item1.ModeId;
                                modeStartEndDateTime.tempModeId = item1.TempModeId;
                                modeStartEndDateTimes.Add(modeStartEndDateTime);
                            }
                            obj.isStatus = true;
                            obj.response = modeStartEndDateTimes;
                            obj.tempModeIds = String.Join(",", modeStartEndDateTimes.OrderBy(m => m.tempModeId).Select(m => m.tempModeId).ToList());
                        }
                        else
                        {
                            prevTempModeId = check[i - 1].TempModeId;

                            var newTempMode = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();

                            var dbCheck = db.TblTempMode.Where(m => m.TempModeId == prevTempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                dbCheck.EndTime = newTempMode.EndTime;
                                DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                var durationInSec = (EndDateTime - StartDateTime).TotalSeconds;
                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                                db.SaveChanges();

                                db.TblTempMode.Remove(newTempMode);
                                db.SaveChanges();

                            }

                            var tempData = db.TblTempMode.Where(m => m.ModeId == dbCheck.ModeId && m.MachineId == dbCheck.MachineId & m.IsSaved == 0).ToList();
                            List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();
                            foreach (var item1 in tempData)
                            {
                                DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
                                DateTime StartDateTime = Convert.ToDateTime(item1.StartTime);

                                string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids = dt.Split();
                                string endDate = ids[0];
                                string endTime = ids[1];

                                string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                string[] ids1 = dt1.Split();
                                string startDate = ids1[0];
                                string startTime = ids1[1];

                                ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                                modeStartEndDateTime.startDate = startDate;
                                modeStartEndDateTime.startTime = startTime;
                                modeStartEndDateTime.endDate = endDate;
                                modeStartEndDateTime.endTime = endTime;
                                modeStartEndDateTime.modeId = item1.ModeId;
                                modeStartEndDateTime.tempModeId = item1.TempModeId;
                                modeStartEndDateTimes.Add(modeStartEndDateTime);
                            }
                            obj.isStatus = true;
                            obj.response = modeStartEndDateTimes;
                            obj.tempModeIds = String.Join(",", modeStartEndDateTimes.OrderBy(m => m.tempModeId).Select(m => m.tempModeId).ToList());
                        }
                    }
                    //else
                    //{
                    //    int AfterTempModeId = check[i + 1].TempModeId;
                    //    var newTempMode = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();

                    //    var dbCheck = db.TblTempMode.Where(m => m.TempModeId == AfterTempModeId && m.ModeId == modeId && m.IsApproved == 0 && m.IsSaved == 0).FirstOrDefault();
                    //    if (dbCheck != null)
                    //    {
                    //        dbCheck.EndTime = newTempMode.EndTime;
                    //        dbCheck.StartTime = newTempMode.StartTime;
                    //        db.SaveChanges();

                    //        db.TblTempMode.Remove(newTempMode);
                    //        db.SaveChanges();

                    //    }

                    //    var tempData = db.TblTempMode.Where(m => m.ModeId == dbCheck.ModeId && m.MachineId == dbCheck.MachineId && m.IsSaved == 0).ToList();
                    //    List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();
                    //    foreach (var item1 in tempData)
                    //    {
                    //        DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
                    //        DateTime StartDateTime = Convert.ToDateTime(item1.StartTime);

                    //        string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    //        string[] ids = dt.Split();
                    //        string endDate = ids[0];
                    //        string endTime = ids[1];

                    //        string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    //        string[] ids1 = dt1.Split();
                    //        string startDate = ids1[0];
                    //        string startTime = ids1[1];

                    //        ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                    //        modeStartEndDateTime.startDate = startDate;
                    //        modeStartEndDateTime.startTime = startTime;
                    //        modeStartEndDateTime.endDate = endDate;
                    //        modeStartEndDateTime.endTime = endTime;
                    //        modeStartEndDateTime.modeId = item1.ModeId;
                    //        modeStartEndDateTime.tempModeId = item1.TempModeId;
                    //        modeStartEndDateTimes.Add(modeStartEndDateTime);
                    //    }
                    //    obj.isStatus = true;
                    //    obj.response = modeStartEndDateTimes;
                    //}
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Delete Temp Table Data when model closes
        /// </summary>
        /// <param name="modeId"></param>
        /// <param name="machineId"></param>
        /// <returns></returns>
        public CommonResponse1 DeleteTempTableData(string tempmodeIds)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                //var check = db.TblTempMode.Where(m => m.MachineId == machineId && m.ModeId == modeId).OrderBy(m => m.TempModeId).ToList();
                //if (check.Count > 0)
                //{
                //    for (int i = 0; i < check.Count; i++)
                //    {
                //        if (i == 0)
                //        {
                //            check[i].EndTime = check[check.Count - 1].EndTime;
                //            db.SaveChanges();
                //        }
                //        else
                //        {
                //            db.TblTempMode.Remove(check[i]);
                //            db.SaveChanges();
                //        }

                //    }
                //    obj.isStatus = true;
                //    obj.response = "Deleted Successfully";
                string[] aary = tempmodeIds.Split(',');
                List<int> intArry = aary.ToList().Select(int.Parse).ToList();
                var check = db.TblTempMode.Where(m => intArry.Contains(m.TempModeId)).ToList();

                if (check.Count > 0)
                {
                    for (int i = 0; i < check.Count; i++)
                    {
                        if (i == 0)
                        {
                            check[i].EndTime = check[check.Count - 1].EndTime;
                            db.SaveChanges();
                        }
                        else
                        {
                            db.TblTempMode.Remove(check[i]);
                            db.SaveChanges();
                        }
                    }
                }
                obj.isStatus = true;
                obj.response = "Deleted Successfully";
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// After saving the mode details to temp table
        /// </summary>
        /// <param name="tempModeIds"></param>
        /// <returns></returns>
        public CommonResponse1 AddModeDetails(string tempModeIds)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                string[] ids = tempModeIds.Split(',');
                foreach (var item in ids)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                    if (check != null)
                    {
                        check.IsSaved = 1;
                        db.SaveChanges();
                        obj.isStatus = true;
                        obj.response = "Added Successfully";
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

        /// <summary>
        /// Send Mail To Approve Or Reject
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 SendMailToApproveOrReject(SendMailDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                #region
                string CorrectedDate = "";
                try
                {
                    string[] dt = data.date.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    CorrectedDate = frDate;
                }
                catch
                {
                    CorrectedDate = data.date;
                }
                #endregion

                int sl = 1;
                string MachineName = db.Tblmachinedetails.Where(m => m.MachineId == data.machineId).Select(m => m.MachineInvNo).FirstOrDefault();
                string PlantName = db.Tblplant.Where(m => m.PlantId == data.plantId).Select(m => m.PlantName).FirstOrDefault();
                string ShopName = db.Tblshop.Where(m => m.ShopId == data.shopId).Select(m => m.ShopName).FirstOrDefault();
                string cellName = db.Tblcell.Where(m => m.CellId == data.cellId).Select(m => m.CellName).FirstOrDefault();
                int machineId = data.machineId;

                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\DataLossTcfTemplate1.html");
                string htmlStr = File.ReadAllText(reader);


                string[] ids = data.tempModeIds.Split(',');

                string mainTemp = "<tr><td><span>{{slno}}</span> </td> <td><span>{{machinename}}</span> </td><td><span>{{mode}}" +
                    "</span> </td><td><span>{{rl1}}</span></td><td><span>{{rl2}}</span></td><td><span>{{rl3}}</span></td><td><span>{{startTime}}</span></td><td><span>{{endTime}}</span></td></tr>";
                string strReplace = "";
                //int appLevl = 0;
                //bool sendMail = false;
                foreach (var item in ids)
                {
                    int tempModeID = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 0 && m.OverAllSaved == 1).FirstOrDefault();
                    if (check != null)
                    {
                        check.IsApproved = 1; //send for approval                  
                        check.MailSendDate = DateTime.Now;
                        check.IsHold = 1;
                        db.SaveChanges();

                        var lossData = (from wf in db.TblTempLiveLossOfEntry
                                        join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                        where wf.TempModeId == tempModeID && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == CorrectedDate
                                        select new
                                        {
                                            LossCodeId = lc.LossCodeId,
                                            LossCode = lc.LossCode,
                                            LossCodesLevel = lc.LossCodesLevel,
                                            LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                            LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                            StartDateTime = wf.StartDateTime,
                                            EndDateTime = wf.EndDateTime,
                                            MachineId = wf.MachineId,
                                            MessageCodeId = wf.MessageCodeId,
                                            CorrectedDate = wf.CorrectedDate,
                                            TempLossId = wf.TempLossId,

                                            ress1 = (lc.LossCodesLevel == 1 ?
                                             lc.LossCode :
                                             db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                            ress2 = (lc.LossCodesLevel == 2 ?
                                             lc.LossCode :
                                             db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                            ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                        }).ToList();

                        if (lossData.Count > 0)
                        {
                            string temp = mainTemp;
                            string slno = Convert.ToString(sl);
                            foreach (var res in lossData)
                            {
                                temp = temp.Replace("{{slno}}", slno);
                                temp = temp.Replace("{{machinename}}", MachineName);
                                temp = temp.Replace("{{mode}}", check.Mode);
                                temp = temp.Replace("{{rl1}}", res.ress1);
                                temp = temp.Replace("{{rl2}}", res.ress2);
                                temp = temp.Replace("{{rl3}}", res.ress3);
                                temp = temp.Replace("{{startTime}}", check.StartTime.ToString());
                                temp = temp.Replace("{{endTime}}", check.EndTime.ToString());
                                strReplace = strReplace + temp;
                            }
                        }
                        else
                        {
                            string temp = mainTemp;
                            string slno = Convert.ToString(sl);
                            temp = temp.Replace("{{slno}}", slno);
                            temp = temp.Replace("{{machinename}}", MachineName);
                            temp = temp.Replace("{{mode}}", check.Mode);
                            temp = temp.Replace("{{rl1}}", "-");
                            temp = temp.Replace("{{rl2}}", "-");
                            temp = temp.Replace("{{rl3}}", "-");
                            temp = temp.Replace("{{startTime}}", check.StartTime.ToString());
                            temp = temp.Replace("{{endTime}}", check.EndTime.ToString());
                            strReplace = strReplace + temp;
                        }
                        sl++;
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "Already Mail Sent";
                    }
                }

                htmlStr = htmlStr.Replace("{{strReplace}}", strReplace);
                htmlStr = htmlStr.Replace("{{LossLevel}}", "1");
                //string acceptUrl = configuration.GetSection("AppSettings").GetSection("AcceptURLDASDataCorr").Value;
                //string rejectUrl = configuration.GetSection("AppSettings").GetSection("RejectURLDASDataCorr").Value;

                //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLDASDataCorr").Value;
                string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLDASDataCorr").Value;

                string rejectSrc = rejectUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + data.tempModeIds + "&plantID=" + data.plantId + "&cellId=" + data.cellId + "&shopId=" + data.shopId + "&checked=0";
                //string acceptSrc = acceptUrl + "CorrectedDate=" + CorrectedDate + "&machineId=" + machineId + "&tempId=" + data.tempModeIds + "&plantID=" + data.plantId + "&cellId=" + data.cellId + "&shopId=" + data.shopId + "";

                string FirstApproverTomailIds = " ";
                string FirstApproverCcmailIds = " ";

                var mailDetails = (from wf in db.TblTcfApprovedMaster
                                   where wf.TcfModuleId == 3 && (wf.CellId == data.cellId || wf.ShopId == data.shopId)
                                   select new
                                   {
                                       FirstApproverToList = wf.FirstApproverToList,
                                       FirstApproverCcList = wf.FirstApproverCcList,
                                       SecondApproverToList = wf.SecondApproverToList,
                                       SecondApproverCcList = wf.SecondApproverCcList
                                   }).ToList();
                if (mailDetails.Count > 0)
                {
                    foreach (var items in mailDetails)
                    {
                        if (items.FirstApproverToList != null && items.FirstApproverCcList != null)
                        {
                            FirstApproverTomailIds = items.FirstApproverToList;
                            FirstApproverCcmailIds = items.FirstApproverCcList;
                        }
                    }
                }

                //htmlStr = htmlStr.Replace("{{urlAR}}", acceptSrc);
                htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);


                bool ret = SendMail(htmlStr, FirstApproverTomailIds, FirstApproverCcmailIds, 1, MachineName, CorrectedDate, null);
                if (ret)
                {
                    obj.isStatus = true;
                    obj.response = "Mail Sent Successfully";
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "Mail Not Sent";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// For Sending mail
        /// </summary>
        /// <param name="message"></param>
        /// <param name="toList"></param>
        /// <param name="ccList"></param>
        /// <param name="image"></param>
        /// <returns></returns>
        public bool SendMail(string message, string toList, string ccList, int image, string machineName, string correctedDate, string sendMessage)
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

                    var smtpConn = db.Smtpdetails.Where(x => x.IsDeleted == true && x.TcfModuleId == 3).FirstOrDefault();
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
                    //mail.From = new MailAddress("monika.ms@srkssolutions.com");
                    mail.Subject = "Data Loss TCF" + " " + machineName + " " + correctedDate;
                    if (sendMessage == null)
                    {
                        mail.Body = "" + message;
                    }
                    else if (sendMessage != null)
                    {
                        mail.Body = "" + message + "" + sendMessage;
                    }
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

        /// <summary>
        /// Get Reject Reasons
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetRejectReasons()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var rejectReassondata = (from wf in db.Tblrejectreason
                                         where wf.IsTcf == 1 && wf.IsDeleted == 0
                                         select new
                                         {
                                             rejectName = wf.RejectName,
                                             reasonId = wf.Rid
                                         }).ToList();
                if (rejectReassondata.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = rejectReassondata;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Send Mail After Reject
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 SendMailAfterReject(SendMailDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                #region
                string CorrectedDate = "";
                try
                {
                    string[] dt = data.date.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    CorrectedDate = frDate;
                }
                catch
                {
                    CorrectedDate = data.date;
                }
                #endregion

                #region
                string[] tempModeIds = data.tempModeIds.Split(',');
                string machineName = "";
                string TomailIds = " ";
                string CcmailIds = " ";
                string SecondApproverToMailIds = " ";
                string SecondApproverCcMailIds = " ";


                foreach (var item in tempModeIds)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsApproved == 1 && m.ApproveLevel == null).FirstOrDefault();
                    if (check != null)
                    {
                        check.IsApproved = 3;
                        check.RejectLevel = 1;
                        check.RejectReasonId = data.reasonId;
                        db.SaveChanges();

                        var mailDetails = (from wf in db.TblTcfApprovedMaster
                                           where wf.TcfModuleId == 3 && (wf.CellId == data.cellId || wf.ShopId == data.shopId)
                                           select new
                                           {
                                               FirstApproverToList = wf.FirstApproverToList,
                                               FirstApproverCcList = wf.FirstApproverCcList
                                           }).ToList();
                        if (mailDetails.Count > 0)
                        {
                            foreach (var items in mailDetails)
                            {
                                TomailIds = items.FirstApproverToList;
                                CcmailIds = items.FirstApproverCcList;
                            }
                        }
                    }

                    var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsApproved == 2 && m.ApproveLevel == 1).FirstOrDefault();
                    if (dbCheck != null)
                    {
                        dbCheck.IsApproved = 3;
                        dbCheck.RejectLevel = 2;
                        dbCheck.RejectReasonId = data.reasonId;
                        db.SaveChanges();

                        var mailDetails1 = (from wf in db.TblTcfApprovedMaster
                                            where wf.TcfModuleId == 3 && (wf.CellId == data.cellId || wf.ShopId == data.shopId)
                                            select new
                                            {
                                                SecondApproverToList = wf.SecondApproverToList,
                                                SecondApproverCcList = wf.SecondApproverCcList
                                            }).ToList();

                        foreach (var items in mailDetails1)
                        {
                            if (items.SecondApproverToList != null && items.SecondApproverCcList != null)
                            {
                                SecondApproverToMailIds = items.SecondApproverToList;
                                SecondApproverCcMailIds = items.SecondApproverCcList;
                            }
                        }
                    }

                    machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                }

                var mode = db.TblModeBackUp.Where(m => m.CorrectedDate == CorrectedDate && m.MachineId == data.machineId).ToList();

                foreach (var items in mode)
                {
                    #region add new row in live mode table which we are going to insert in mode table
                    Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                    tbllivemodedb.MachineId = Convert.ToInt32(items.MachineId);
                    tbllivemodedb.Mode = items.Mode;
                    tbllivemodedb.ColorCode = items.ColorCode;
                    tbllivemodedb.InsertedOn = DateTime.Now;
                    tbllivemodedb.InsertedBy = 1;
                    tbllivemodedb.CorrectedDate = items.CorrectedDate;
                    tbllivemodedb.IsDeleted = 0;
                    tbllivemodedb.StartTime = items.StartTime;
                    tbllivemodedb.EndTime = items.EndTime;
                    tbllivemodedb.IsCompleted = Convert.ToInt32(items.IsCompleted);
                    tbllivemodedb.DurationInSec = items.DurationInSec;
                    db.Tbllivemodedb.Add(tbllivemodedb);
                    db.SaveChanges();
                    #endregion

                    #region add new row in mode table which we have inserted in live mode table in last region
                    Tblmode tblmode = new Tblmode();
                    tblmode.MachineId = Convert.ToInt32(items.MachineId);
                    tblmode.ColorCode = tbllivemodedb.ColorCode;
                    tblmode.ModeId = tbllivemodedb.ModeId;
                    tblmode.Mode = tbllivemodedb.Mode;
                    tblmode.IsDeleted = 0;
                    tblmode.InsertedOn = DateTime.Now;
                    tblmode.InsertedBy = 1;
                    tblmode.CorrectedDate = items.CorrectedDate;
                    tblmode.StartTime = tbllivemodedb.StartTime;
                    tblmode.EndTime = tbllivemodedb.EndTime;
                    tblmode.IsCompleted = 1;
                    tblmode.DurationInSec = tbllivemodedb.DurationInSec;
                    db.Tblmode.Add(tblmode);
                    db.SaveChanges();
                    #endregion
                }

                string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml'><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
               "</head><body><p>Dear,</p></br><p><center> The Data Loss Has Been Rejected</center></p></br><p>Thank you" +
               "</p></br><p>Sincerely,</p></br></body></html>";

                bool ret = false;

                if ((TomailIds != null || TomailIds != "") && (CcmailIds != null || CcmailIds != ""))
                {
                    ret = SendMail(message, TomailIds, CcmailIds, 1, machineName, CorrectedDate, null);
                }
                if ((SecondApproverToMailIds != null || SecondApproverToMailIds != "") && (SecondApproverCcMailIds != null || SecondApproverCcMailIds != ""))
                {
                    ret = SendMail(message, SecondApproverToMailIds, SecondApproverCcMailIds, 1, machineName, CorrectedDate, null);
                }

                if (ret)
                {
                    obj.isStatus = true;
                    obj.response = "Mail Sent Successfully";
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "Message Not Sent";
                }
                #endregion
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// After Split Duration
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 AfterSplitDuration(int tempModeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            List<ModeStartEndDateTimeWithReasons> modeStartEndDateTimeWithReasonsList = new List<ModeStartEndDateTimeWithReasons>();
            try
            {
                var check = (from wf in db.TblTempMode
                             where wf.IsDeleted == 0 && wf.TempModeId == tempModeId && wf.IsApproved == 0 && wf.OverAllSaved == null
                             select new
                             {
                                 CorrectedDate = wf.CorrectedDate,
                                 MachineId = wf.MachineId
                             }).FirstOrDefault();

                var dbCheck = (from tm in db.TblTempMode
                               where tm.IsDeleted == 0 && tm.MachineId == check.MachineId && tm.CorrectedDate == check.CorrectedDate && tm.IsApproved == 0 && (tm.OverAllSaved == null || tm.OverAllSaved == 1)
                               select new
                               {
                                   StartTime = tm.StartTime,
                                   EndTime = tm.EndTime,
                                   TempModeId = tm.TempModeId,
                                   ModeId = tm.ModeId,
                                   MachineId = tm.MachineId,
                                   CorrectedDate = tm.CorrectedDate,
                                   Mode = tm.Mode,
                                   OverAllSaved = tm.OverAllSaved
                               }).ToList();
                List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();

                if (dbCheck.Count > 0)
                {
                    foreach (var item in dbCheck)
                    {
                        var liveData = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == item.TempModeId).FirstOrDefault();
                        if (liveData != null)
                        {
                            var lossCodes = (from wf in db.TblTempLiveLossOfEntry
                                             join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                             where wf.DoneWithRow == 1 && wf.TempModeId == item.TempModeId && wf.MachineId == item.MachineId && wf.CorrectedDate == item.CorrectedDate
                                             select new
                                             {
                                                 LossCodeId = lc.LossCodeId,
                                                 LossCode = lc.LossCode,
                                                 LossCodesLevel = lc.LossCodesLevel,
                                                 LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                 LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                 StartDateTime = wf.StartDateTime,
                                                 EndDateTime = wf.EndDateTime,
                                                 MachineId = wf.MachineId,
                                                 MessageCodeId = wf.MessageCodeId,
                                                 CorrectedDate = wf.CorrectedDate,
                                                 SplossId = wf.TempLossId,

                                                 ress1 = (lc.LossCodesLevel == 1 ?
                                                     new { lc.LossCode, lc.LossCodeId } :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => new { m.LossCode, m.LossCodeId }).FirstOrDefault()),

                                                 ress2 = (lc.LossCodesLevel == 2 ?
                                                     new { lc.LossCode, lc.LossCodeId } :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => new { m.LossCode, m.LossCodeId }).FirstOrDefault()),

                                                 ress3 = (lc.LossCodesLevel == 3 ? new { lc.LossCode, lc.LossCodeId } : null),

                                                 //res1LossCodeID = (lc.LossCodesLevel == 1 ?
                                                 //    lc.LossCodesLevel1Id :
                                                 //    db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCodeId).FirstOrDefault()),

                                                 //res2LossCodeID = (lc.LossCodesLevel == 2 ?
                                                 //    lc.LossCodesLevel2Id :
                                                 //    db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCodeId).FirstOrDefault()),

                                                 //res3LossCodeID = (lc.LossCodesLevel == 3 ? lc.LossCodeId :
                                                 //    db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodeId).Select(m => m.LossCodeId).FirstOrDefault())

                                             }).ToList();


                            foreach (var mode in lossCodes)
                            {
                                ModeStartEndDateTimeWithReasons modeStartEndDateTimeWithReasons = new ModeStartEndDateTimeWithReasons();
                                DateTime startDateTime = Convert.ToDateTime(item.StartTime);
                                string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids = dt.Split();
                                //modeStartEndDateTimeWithReasons.startDate = ids[0];
                                //modeStartEndDateTimeWithReasons.startTime = ids[1];
                                modeStartEndDateTimeWithReasons.startDate = dt;
                                DateTime EndDateTime = Convert.ToDateTime(item.EndTime);
                                string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids1 = dt1.Split();
                                //modeStartEndDateTimeWithReasons.endDate = ids1[0];
                                //modeStartEndDateTimeWithReasons.endTime = ids1[1];
                                modeStartEndDateTimeWithReasons.endDate = dt1;
                                modeStartEndDateTimeWithReasons.modeId = item.ModeId;
                                modeStartEndDateTimeWithReasons.tempModeId = item.TempModeId;
                                if (mode.ress1 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res1 = mode.ress1.LossCodeId;
                                }
                                if (mode.ress2 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res2 = mode.ress2.LossCodeId;
                                }
                                if (mode.ress3 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res3 = mode.ress3.LossCodeId;
                                }
                                if (mode.ress1 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res1Name = mode.ress1.LossCode;
                                }
                                if (mode.ress2 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res2Name = mode.ress2.LossCode;
                                }
                                if (mode.ress3 != null)
                                {
                                    modeStartEndDateTimeWithReasons.res3Name = mode.ress3.LossCode;
                                }

                                if (item.Mode == "IDLE")
                                {
                                    modeStartEndDateTimeWithReasons.mode = "Idle";
                                }
                                else if (item.Mode == "BREAKDOWN")
                                {
                                    modeStartEndDateTimeWithReasons.mode = "Breakdown";
                                }
                                modeStartEndDateTimeWithReasonsList.Add(modeStartEndDateTimeWithReasons);
                            }
                        }
                        else
                        {
                            ModeStartEndDateTimeWithReasons modeStartEndDateTimeWithReasons = new ModeStartEndDateTimeWithReasons();
                            DateTime startDateTime = Convert.ToDateTime(item.StartTime);
                            string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                            modeStartEndDateTimeWithReasons.startDate = dt;
                            DateTime EndDateTime = Convert.ToDateTime(item.EndTime);
                            string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                            modeStartEndDateTimeWithReasons.endDate = dt1;
                            modeStartEndDateTimeWithReasons.modeId = item.ModeId;
                            modeStartEndDateTimeWithReasons.tempModeId = item.TempModeId;
                            modeStartEndDateTimeWithReasons.mode = item.Mode;
                            modeStartEndDateTimeWithReasons.overAllSaved = Convert.ToInt32(item.OverAllSaved);
                            modeStartEndDateTimeWithReasonsList.Add(modeStartEndDateTimeWithReasons);
                        }
                    }
                    obj.isStatus = true;
                    obj.response = modeStartEndDateTimeWithReasonsList.OrderBy(m => m.startDate).ThenBy(m => m.startTime);
                }

                //var tempMode = (from wf in db.TblTempMode
                //                where wf.ModeId == modeId
                //                select new
                //                {
                //                    CorrectedDate = wf.CorrectedDate,
                //                    MachineId = wf.MachineId

                //                }).FirstOrDefault();

                //var dbCheck = (from wf in db.TblTempMode
                //               where wf.CorrectedDate == tempMode.CorrectedDate && wf.MachineId == tempMode.MachineId && wf.ModeId != modeId
                //               select new
                //               {
                //                   StartTime = wf.StartTime,
                //                   EndTime = wf.EndTime,
                //                   ModeId = wf.ModeId,
                //                   MachineId = wf.MachineId,
                //                   TempModeId = wf.TempModeId

                //               }).ToList();
                //if (dbCheck.Count > 0)
                //{
                //    foreach (var item1 in dbCheck)
                //    {
                //        ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                //        DateTime startDateTime = Convert.ToDateTime(item1.StartTime);
                //        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                //        string[] ids = dt.Split();
                //        modeStartEndDateTime.startDate = ids[0];
                //        modeStartEndDateTime.startTime = ids[1];
                //        DateTime EndDateTime = Convert.ToDateTime(item1.EndTime);
                //        string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                //        string[] ids1 = dt1.Split();
                //        modeStartEndDateTime.endDate = ids1[0];
                //        modeStartEndDateTime.endTime = ids1[1];
                //        modeStartEndDateTime.modeId = item1.ModeId;
                //        modeStartEndDateTime.tempModeId = item1.TempModeId;
                //        modeStartEndDateTimeList.Add(modeStartEndDateTime);
                //        obj.isStatus = true;
                //        obj.response = modeStartEndDateTimeList.OrderBy(m => m.startTime);
                //    }
                //}
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Mode Details
        /// </summary>
        /// <param name="MachineId"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public CommonResponse1 GetModeDetails(int MachineId, string date)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                #region
                string CorrectedDate = "";
                try
                {
                    string[] dt = date.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    CorrectedDate = frDate;
                }
                catch
                {
                    CorrectedDate = date;
                }
                #endregion

                List<ModeStartEndDateTime> modeStartEndDateTimeList = new List<ModeStartEndDateTime>();

                #region get mode details either from temp table or mode table

                var checktemp = db.TblTempMode.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate).ToList();
                if (checktemp.Count > 0)
                {
                    var tempDataDet = db.TblTempMode.Where(wf => wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate && wf.IsApproved == 0).ToList();

                    var tempDataDet1 = db.TblTempMode.Where(wf => wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate && ((wf.IsApproved == 2 && wf.ApproveLevel == 1) || (wf.IsApproved == 2 && wf.ApproveLevel == 2) || wf.IsApproved == 3)).ToList();

                    var tempDataApproval = db.TblTempMode.Where(wf => wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate && wf.IsApproved == 1).ToList();

                    if (tempDataDet.Count > 0)
                    {
                        db.TblTempMode.RemoveRange(tempDataDet);
                        db.SaveChanges();

                        var templive = db.TblTempLiveLossOfEntry.Where(wf => wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate).ToList();
                        db.TblTempLiveLossOfEntry.RemoveRange(templive);
                        db.SaveChanges();

                        var modeBackUp = db.TblModeBackUp.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate).ToList();
                        db.TblModeBackUp.RemoveRange(modeBackUp);
                        db.SaveChanges();

                        var tblMode = db.Tblmode.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate && m.Mode=="PowerOff").ToList();

                        if (tblMode.Count > 0)
                        {
                            foreach (var item in tblMode)
                            {
                                TblModeBackUp tblModeBackUp = new TblModeBackUp();
                                tblModeBackUp.ModeId = item.ModeId;
                                tblModeBackUp.MachineId = item.MachineId;
                                tblModeBackUp.Mode = item.Mode;
                                tblModeBackUp.InsertedOn = DateTime.Now;
                                tblModeBackUp.InsertedBy = item.InsertedBy;
                                tblModeBackUp.ModifiedOn = item.ModifiedOn;
                                tblModeBackUp.ModifiedBy = item.ModifiedBy;
                                tblModeBackUp.CorrectedDate = item.CorrectedDate;
                                tblModeBackUp.IsDeleted = item.IsDeleted;
                                tblModeBackUp.StartTime = item.StartTime;
                                tblModeBackUp.EndTime = item.EndTime;
                                tblModeBackUp.ColorCode = item.ColorCode;
                                tblModeBackUp.IsCompleted = item.IsCompleted;
                                tblModeBackUp.DurationInSec = item.DurationInSec;
                                tblModeBackUp.ModeMonth = item.ModeMonth;
                                tblModeBackUp.ModeQuarter = item.ModeQuarter;
                                tblModeBackUp.ModeWeekNumber = item.ModeWeekNumber;
                                db.TblModeBackUp.Add(tblModeBackUp);
                                db.SaveChanges();

                                TblTempMode tblTempMode = new TblTempMode();
                                tblTempMode.ModeId = item.ModeId;
                                tblTempMode.MachineId = item.MachineId;
                                tblTempMode.Mode = item.Mode;
                                tblTempMode.InsertedOn = DateTime.Now;
                                tblTempMode.InsertedBy = item.InsertedBy;
                                tblTempMode.ModifiedOn = item.ModifiedOn;
                                tblTempMode.ModifiedBy = item.ModifiedBy;
                                tblTempMode.CorrectedDate = item.CorrectedDate;
                                tblTempMode.IsDeleted = item.IsDeleted;
                                tblTempMode.StartTime = item.StartTime;
                                tblTempMode.EndTime = item.EndTime;
                                tblTempMode.ColorCode = item.ColorCode;
                                tblTempMode.IsCompleted = item.IsCompleted;
                                tblTempMode.DurationInSec = item.DurationInSec;
                                tblTempMode.ModeMonth = item.ModeMonth;
                                tblTempMode.ModeQuarter = item.ModeQuarter;
                                tblTempMode.ModeWeekNumber = item.ModeWeekNumber;
                                tblTempMode.IsApproved = 0;
                                tblTempMode.IsSaved = 0;
                                //tblTempMode.UpdateLevel = updateLevel;
                                db.TblTempMode.Add(tblTempMode);
                                db.SaveChanges();

                                ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                                DateTime startDateTime = Convert.ToDateTime(tblTempMode.StartTime);
                                string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids = dt.Split();
                                //modeStartEndDateTime.startDate = ids[0];
                                //modeStartEndDateTime.startTime = ids[1];
                                DateTime EndDateTime = Convert.ToDateTime(tblTempMode.EndTime);
                                string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids1 = dt1.Split();
                                //modeStartEndDateTime.endDate = ids1[0];
                                //modeStartEndDateTime.endTime = ids1[1];
                                modeStartEndDateTime.endDate = dt1;
                                modeStartEndDateTime.startDate = dt;
                                modeStartEndDateTime.modeId = tblTempMode.ModeId;
                                modeStartEndDateTime.tempModeId = tblTempMode.TempModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime);
                            }
                            obj.isStatus = true;
                            obj.response = modeStartEndDateTimeList;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.response = "No Item Found";
                        }
                    }
                    else if (tempDataDet1.Count > 0)
                    {
                        var tblMode = db.Tblmode.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate && m.Mode == "PowerOff").ToList();

                        var modeBackUp = db.TblModeBackUp.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate).ToList();
                        db.TblModeBackUp.RemoveRange(modeBackUp);
                        db.SaveChanges();

                        if (tblMode.Count > 0)
                        {
                            foreach (var item in tblMode)
                            {
                                //int updateLevel = 1;

                                //var machine = (from wf in db.Tblmachinedetails
                                //               where wf.MachineId == MachineId
                                //               select new
                                //               {
                                //                   shopId = wf.ShopId,
                                //                   cellId = wf.CellId,
                                //                   plantId = wf.PlantId
                                //               }).FirstOrDefault();

                                //var getMailIdsLevel = db.TblTcfApprovedMaster.Where(m => m.TcfModuleId == 3 && m.IsDeleted == 0 && m.PlantId == machine.plantId && (m.ShopId == machine.shopId || m.CellId == machine.cellId)).FirstOrDefault();
                                TblModeBackUp tblModeBackUp = new TblModeBackUp();
                                tblModeBackUp.ModeId = item.ModeId;
                                tblModeBackUp.MachineId = item.MachineId;
                                tblModeBackUp.Mode = item.Mode;
                                tblModeBackUp.InsertedOn = DateTime.Now;
                                tblModeBackUp.InsertedBy = item.InsertedBy;
                                tblModeBackUp.ModifiedOn = item.ModifiedOn;
                                tblModeBackUp.ModifiedBy = item.ModifiedBy;
                                tblModeBackUp.CorrectedDate = item.CorrectedDate;
                                tblModeBackUp.IsDeleted = item.IsDeleted;
                                tblModeBackUp.StartTime = item.StartTime;
                                tblModeBackUp.EndTime = item.EndTime;
                                tblModeBackUp.ColorCode = item.ColorCode;
                                tblModeBackUp.IsCompleted = item.IsCompleted;
                                tblModeBackUp.DurationInSec = item.DurationInSec;
                                tblModeBackUp.ModeMonth = item.ModeMonth;
                                tblModeBackUp.ModeQuarter = item.ModeQuarter;
                                tblModeBackUp.ModeWeekNumber = item.ModeWeekNumber;
                                db.TblModeBackUp.Add(tblModeBackUp);
                                db.SaveChanges();

                                TblTempMode tblTempMode = new TblTempMode();
                                tblTempMode.ModeId = item.ModeId;
                                tblTempMode.MachineId = item.MachineId;
                                tblTempMode.Mode = item.Mode;
                                tblTempMode.InsertedOn = DateTime.Now;
                                tblTempMode.InsertedBy = item.InsertedBy;
                                tblTempMode.ModifiedOn = item.ModifiedOn;
                                tblTempMode.ModifiedBy = item.ModifiedBy;
                                tblTempMode.CorrectedDate = item.CorrectedDate;
                                tblTempMode.IsDeleted = item.IsDeleted;
                                tblTempMode.StartTime = item.StartTime;
                                tblTempMode.EndTime = item.EndTime;
                                tblTempMode.ColorCode = item.ColorCode;
                                tblTempMode.IsCompleted = item.IsCompleted;
                                tblTempMode.DurationInSec = item.DurationInSec;
                                tblTempMode.ModeMonth = item.ModeMonth;
                                tblTempMode.ModeQuarter = item.ModeQuarter;
                                tblTempMode.ModeWeekNumber = item.ModeWeekNumber;
                                tblTempMode.IsApproved = 0;
                                tblTempMode.IsSaved = 0;
                                //tblTempMode.UpdateLevel = updateLevel;
                                db.TblTempMode.Add(tblTempMode);
                                db.SaveChanges();

                                ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                                DateTime startDateTime = Convert.ToDateTime(tblTempMode.StartTime);
                                string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids = dt.Split();
                                //modeStartEndDateTime.startDate = ids[0];
                                //modeStartEndDateTime.startTime = ids[1];
                                DateTime EndDateTime = Convert.ToDateTime(tblTempMode.EndTime);
                                string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                //string[] ids1 = dt1.Split();
                                //modeStartEndDateTime.endDate = ids1[0];
                                //modeStartEndDateTime.endTime = ids1[1];
                                modeStartEndDateTime.endDate = dt1;
                                modeStartEndDateTime.startDate = dt;
                                modeStartEndDateTime.modeId = tblTempMode.ModeId;
                                modeStartEndDateTime.tempModeId = tblTempMode.TempModeId;
                                modeStartEndDateTimeList.Add(modeStartEndDateTime);
                            }
                            obj.isStatus = true;
                            obj.response = modeStartEndDateTimeList;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.response = "No Item Found";
                        }
                    }
                    else if (tempDataApproval.Count > 0)
                    {
                        obj.isStatus = false;
                        obj.response = "Already Send To Approval";
                    }
                }
                else
                {
                    var tblMode = db.Tblmode.Where(m => m.MachineId == MachineId && m.CorrectedDate == CorrectedDate && m.Mode == "PowerOff").ToList();
                    if (tblMode.Count > 0)
                    {
                        foreach (var item in tblMode)
                        {
                            //int updateLevel = 1;

                            //var machine = (from wf in db.Tblmachinedetails
                            //               where wf.MachineId == MachineId
                            //               select new
                            //               {
                            //                   shopId = wf.ShopId,
                            //                   cellId = wf.CellId,
                            //                   plantId = wf.PlantId
                            //               }).FirstOrDefault();

                            //var getMailIdsLevel = db.TblTcfApprovedMaster.Where(m => m.TcfModuleId == 3 && m.IsDeleted == 0 && m.PlantId == machine.plantId && (m.ShopId == machine.shopId || m.CellId == machine.cellId)).FirstOrDefault();
                            TblModeBackUp tblModeBackUp = new TblModeBackUp();
                            tblModeBackUp.ModeId = item.ModeId;
                            tblModeBackUp.MachineId = item.MachineId;
                            tblModeBackUp.Mode = item.Mode;
                            tblModeBackUp.InsertedOn = DateTime.Now;
                            tblModeBackUp.InsertedBy = item.InsertedBy;
                            tblModeBackUp.ModifiedOn = item.ModifiedOn;
                            tblModeBackUp.ModifiedBy = item.ModifiedBy;
                            tblModeBackUp.CorrectedDate = item.CorrectedDate;
                            tblModeBackUp.IsDeleted = item.IsDeleted;
                            tblModeBackUp.StartTime = item.StartTime;
                            tblModeBackUp.EndTime = item.EndTime;
                            tblModeBackUp.ColorCode = item.ColorCode;
                            tblModeBackUp.IsCompleted = item.IsCompleted;
                            tblModeBackUp.DurationInSec = item.DurationInSec;
                            tblModeBackUp.ModeMonth = item.ModeMonth;
                            tblModeBackUp.ModeQuarter = item.ModeQuarter;
                            tblModeBackUp.ModeWeekNumber = item.ModeWeekNumber;
                            db.TblModeBackUp.Add(tblModeBackUp);
                            db.SaveChanges();

                            TblTempMode tblTempMode = new TblTempMode();
                            tblTempMode.ModeId = item.ModeId;
                            tblTempMode.MachineId = item.MachineId;
                            tblTempMode.Mode = item.Mode;
                            tblTempMode.InsertedOn = DateTime.Now;
                            tblTempMode.InsertedBy = item.InsertedBy;
                            tblTempMode.ModifiedOn = item.ModifiedOn;
                            tblTempMode.ModifiedBy = item.ModifiedBy;
                            tblTempMode.CorrectedDate = item.CorrectedDate;
                            tblTempMode.IsDeleted = item.IsDeleted;
                            tblTempMode.StartTime = item.StartTime;
                            tblTempMode.EndTime = item.EndTime;
                            tblTempMode.ColorCode = item.ColorCode;
                            tblTempMode.IsCompleted = item.IsCompleted;
                            tblTempMode.DurationInSec = item.DurationInSec;
                            tblTempMode.ModeMonth = item.ModeMonth;
                            tblTempMode.ModeQuarter = item.ModeQuarter;
                            tblTempMode.ModeWeekNumber = item.ModeWeekNumber;
                            tblTempMode.IsApproved = 0;
                            tblTempMode.IsSaved = 0;
                            //tblTempMode.UpdateLevel = updateLevel;
                            db.TblTempMode.Add(tblTempMode);
                            db.SaveChanges();

                            ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                            DateTime startDateTime = Convert.ToDateTime(tblTempMode.StartTime);
                            string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                            //string[] ids = dt.Split();
                            //modeStartEndDateTime.startDate = ids[0];
                            //modeStartEndDateTime.startTime = ids[1];
                            DateTime EndDateTime = Convert.ToDateTime(tblTempMode.EndTime);
                            string dt1 = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                            //string[] ids1 = dt1.Split();
                            //modeStartEndDateTime.endDate = ids1[0];
                            //modeStartEndDateTime.endTime = ids1[1];
                            modeStartEndDateTime.endDate = dt1;
                            modeStartEndDateTime.startDate = dt;
                            modeStartEndDateTime.modeId = tblTempMode.ModeId;
                            modeStartEndDateTime.tempModeId = tblTempMode.TempModeId;
                            modeStartEndDateTimeList.Add(modeStartEndDateTime);
                        }
                        obj.isStatus = true;
                        obj.response = modeStartEndDateTimeList;
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "No Item Found";
                    }
                }
                #endregion 
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Approved Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        //public CommonResponse1 GetApprovedDetails(SendMailDetails data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    List<ApproveReject> approveRejectsList = new List<ApproveReject>();
        //    try
        //    {
        //        #region
        //        string CorrectedDate = "";
        //        try
        //        {
        //            string[] dt = data.date.Split('/');
        //            string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            CorrectedDate = frDate;
        //        }
        //        catch
        //        {
        //            CorrectedDate = data.date;
        //        }
        //        #endregion

        //        string[] ids = data.tempModeIds.Split(',');
        //        foreach (var item in ids)
        //        {
        //            int tempModeID = Convert.ToInt32(item);
        //            var check = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 1).FirstOrDefault();
        //            //var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 2 && (m.ApproveLevel == 1 || m.ApproveLevel == 2)).FirstOrDefault();
        //            var dbCheck1 = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 2 && m.ApproveLevel == 1).FirstOrDefault();
        //            //var dbCheck2 = db.TblTempMode.Where(m => m.TempModeId == tempModeID && m.IsApproved == 2 && m.ApproveLevel == 2).FirstOrDefault();
        //            if (check != null)
        //            {
        //                var lossData = (from wf in db.TblTempLiveLossOfEntry
        //                                join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
        //                                where wf.TempModeId == tempModeID && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == CorrectedDate
        //                                select new
        //                                {
        //                                    LossCodeId = lc.LossCodeId,
        //                                    LossCode = lc.LossCode,
        //                                    LossCodesLevel = lc.LossCodesLevel,
        //                                    LossCodesLevel1Id = lc.LossCodesLevel1Id,
        //                                    LossCodesLevel2Id = lc.LossCodesLevel2Id,
        //                                    StartDateTime = wf.StartDateTime,
        //                                    EndDateTime = wf.EndDateTime,
        //                                    MachineId = wf.MachineId,
        //                                    MessageCodeId = wf.MessageCodeId,
        //                                    CorrectedDate = wf.CorrectedDate,
        //                                    TempLossId = wf.TempLossId,

        //                                    ress1 = (lc.LossCodesLevel == 1 ?
        //                                     lc.LossCode :
        //                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


        //                                    ress2 = (lc.LossCodesLevel == 2 ?
        //                                     lc.LossCode :
        //                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

        //                                    ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
        //                                }).ToList();

        //                if (lossData.Count > 0)
        //                {
        //                    foreach (var items in lossData)
        //                    {
        //                        ApproveReject approveReject = new ApproveReject();
        //                        approveReject.reason1 = items.ress1;
        //                        approveReject.reason2 = items.ress2;
        //                        approveReject.reason3 = items.ress3;
        //                        approveReject.machineId = check.MachineId;
        //                        approveReject.mode = check.Mode;
        //                        approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == check.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //                        DateTime startDateTime = Convert.ToDateTime(check.StartTime);
        //                        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                        approveReject.startTime = dt;
        //                        DateTime endDateTime = Convert.ToDateTime(check.EndTime);
        //                        string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                        approveReject.endTime = dt1;
        //                        approveReject.tempModeId = check.TempModeId;
        //                        approveRejectsList.Add(approveReject);
        //                    }
        //                }
        //                else
        //                {
        //                    ApproveReject approveReject = new ApproveReject();
        //                    approveReject.machineId = check.MachineId;
        //                    approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == check.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //                    approveReject.mode = check.Mode;
        //                    //approveReject.startTime = check.StartTime.ToString();
        //                    //approveReject.endTime = check.EndTime.ToString();
        //                    DateTime startDateTime = Convert.ToDateTime(check.StartTime);
        //                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    approveReject.startTime = dt;
        //                    DateTime endDateTime = Convert.ToDateTime(check.EndTime);
        //                    string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    approveReject.endTime = dt1;
        //                    approveReject.tempModeId = check.TempModeId;
        //                    approveRejectsList.Add(approveReject);
        //                }
        //            }
        //            else
        //            {
        //                obj.response = "No ITems Found";
        //                obj.isStatus = false;
        //            }

        //            //if (dbCheck2 != null)
        //            //{
        //            //    var lossData = (from wf in db.TblTempLiveLossOfEntry
        //            //                    join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
        //            //                    where wf.TempModeId == tempModeID && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == CorrectedDate
        //            //                    select new
        //            //                    {
        //            //                        LossCodeId = lc.LossCodeId,
        //            //                        LossCode = lc.LossCode,
        //            //                        LossCodesLevel = lc.LossCodesLevel,
        //            //                        LossCodesLevel1Id = lc.LossCodesLevel1Id,
        //            //                        LossCodesLevel2Id = lc.LossCodesLevel2Id,
        //            //                        StartDateTime = wf.StartDateTime,
        //            //                        EndDateTime = wf.EndDateTime,
        //            //                        MachineId = wf.MachineId,
        //            //                        MessageCodeId = wf.MessageCodeId,
        //            //                        CorrectedDate = wf.CorrectedDate,
        //            //                        TempLossId = wf.TempLossId,

        //            //                        ress1 = (lc.LossCodesLevel == 1 ?
        //            //                         lc.LossCode :
        //            //                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


        //            //                        ress2 = (lc.LossCodesLevel == 2 ?
        //            //                         lc.LossCode :
        //            //                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

        //            //                        ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
        //            //                    }).ToList();

        //            //    if (lossData.Count > 0)
        //            //    {
        //            //        foreach (var items in lossData)
        //            //        {
        //            //            ApproveReject approveReject = new ApproveReject();
        //            //            approveReject.reason1 = items.ress1;
        //            //            approveReject.reason2 = items.ress2;
        //            //            approveReject.reason3 = items.ress3;
        //            //            approveReject.machineId = dbCheck2.MachineId;
        //            //            approveReject.mode = dbCheck2.Mode;
        //            //            approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck2.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //            //            //approveReject.startTime = dbCheck2.StartTime.ToString();
        //            //            //approveReject.endTime = dbCheck2.EndTime.ToString();
        //            //            DateTime startDateTime = Convert.ToDateTime(dbCheck2.StartTime);
        //            //            string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //            //            approveReject.startTime = dt;
        //            //            DateTime endDateTime = Convert.ToDateTime(dbCheck2.EndTime);
        //            //            string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //            //            approveReject.endTime = dt1;
        //            //            approveReject.tempModeId = dbCheck2.TempModeId;
        //            //            approveRejectsList.Add(approveReject);
        //            //        }
        //            //    }
        //            //    else
        //            //    {
        //            //        ApproveReject approveReject = new ApproveReject();
        //            //        approveReject.machineId = dbCheck2.MachineId;
        //            //        approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck2.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //            //        approveReject.mode = dbCheck2.Mode;
        //            //        //approveReject.startTime = dbCheck2.StartTime.ToString();
        //            //        //approveReject.endTime = dbCheck2.EndTime.ToString();
        //            //        DateTime startDateTime = Convert.ToDateTime(dbCheck2.StartTime);
        //            //        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //            //        approveReject.startTime = dt;
        //            //        DateTime endDateTime = Convert.ToDateTime(dbCheck2.EndTime);
        //            //        string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //            //        approveReject.endTime = dt1;
        //            //        approveReject.tempModeId = dbCheck2.TempModeId;
        //            //        approveRejectsList.Add(approveReject);
        //            //    }
        //            //}
        //            //else
        //            //{
        //            //    obj.isStatus = false;
        //            //    obj.response = "No Items Found";
        //            //}

        //            if (dbCheck1 != null)
        //            {
        //                var lossData = (from wf in db.TblTempLiveLossOfEntry
        //                                join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
        //                                where wf.TempModeId == tempModeID && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == CorrectedDate
        //                                select new
        //                                {
        //                                    LossCodeId = lc.LossCodeId,
        //                                    LossCode = lc.LossCode,
        //                                    LossCodesLevel = lc.LossCodesLevel,
        //                                    LossCodesLevel1Id = lc.LossCodesLevel1Id,
        //                                    LossCodesLevel2Id = lc.LossCodesLevel2Id,
        //                                    StartDateTime = wf.StartDateTime,
        //                                    EndDateTime = wf.EndDateTime,
        //                                    MachineId = wf.MachineId,
        //                                    MessageCodeId = wf.MessageCodeId,
        //                                    CorrectedDate = wf.CorrectedDate,
        //                                    TempLossId = wf.TempLossId,

        //                                    ress1 = (lc.LossCodesLevel == 1 ?
        //                                     lc.LossCode :
        //                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


        //                                    ress2 = (lc.LossCodesLevel == 2 ?
        //                                     lc.LossCode :
        //                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

        //                                    ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
        //                                }).ToList();

        //                if (lossData.Count > 0)
        //                {
        //                    foreach (var items in lossData)
        //                    {
        //                        ApproveReject approveReject = new ApproveReject();
        //                        approveReject.reason1 = items.ress1;
        //                        approveReject.reason2 = items.ress2;
        //                        approveReject.reason3 = items.ress3;
        //                        approveReject.machineId = dbCheck1.MachineId;
        //                        approveReject.mode = dbCheck1.Mode;
        //                        approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck1.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //                        //approveReject.startTime = dbCheck1.StartTime.ToString();
        //                        //approveReject.endTime = dbCheck1.EndTime.ToString();
        //                        DateTime startDateTime = Convert.ToDateTime(dbCheck1.StartTime);
        //                        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                        approveReject.startTime = dt;
        //                        DateTime endDateTime = Convert.ToDateTime(dbCheck1.EndTime);
        //                        string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                        approveReject.endTime = dt1;
        //                        approveReject.tempModeId = dbCheck1.TempModeId;
        //                        approveRejectsList.Add(approveReject);
        //                    }
        //                }
        //                else
        //                {
        //                    ApproveReject approveReject = new ApproveReject();
        //                    approveReject.machineId = dbCheck1.MachineId;
        //                    approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck1.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
        //                    approveReject.mode = dbCheck1.Mode;
        //                    //approveReject.startTime = dbCheck1.StartTime.ToString();
        //                    //approveReject.endTime = dbCheck1.EndTime.ToString();
        //                    DateTime startDateTime = Convert.ToDateTime(dbCheck1.StartTime);
        //                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    approveReject.startTime = dt;
        //                    DateTime endDateTime = Convert.ToDateTime(dbCheck1.EndTime);
        //                    string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
        //                    approveReject.endTime = dt1;
        //                    approveReject.tempModeId = dbCheck1.TempModeId;
        //                    approveRejectsList.Add(approveReject);
        //                }
        //            }
        //            else
        //            {
        //                obj.isStatus = false;
        //                obj.response = "No Items Found";
        //            }
        //        }
        //        obj.isStatus = true;
        //        obj.response = approveRejectsList;

        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        /// <summary>
        /// Get Reason Names
        /// </summary>
        /// <param name="lossId"></param>
        /// <returns></returns>
        public List<reasonDetails> GetReasonNames(int lossId)
        {
            //string[] reaons = { };
            //string reasonName = "";
            List<reasonDetails> reasonDetailsList = new List<reasonDetails>();
            try
            {
                var reasonLevel = db.Tbllossescodes.Where(x => x.LossCodeId == lossId).FirstOrDefault();
                if (reasonLevel.LossCodesLevel == 3)
                {
                    reasonDetails reasonDetails = new reasonDetails();
                    reasonDetails.reasonLevel3 = reasonLevel.LossCode;
                    //reasonName = reasonLevel.LossCode;
                    //reaons[0] = reasonName;
                    var reasonName = db.Tbllossescodes.Where(x => x.LossCodeId == reasonLevel.LossCodesLevel2Id).Select(x => x.LossCode).FirstOrDefault();
                    //reaons[1] = reasonName;
                    reasonDetails.reasonLevel2 = reasonName;
                    reasonName = db.Tbllossescodes.Where(x => x.LossCodeId == reasonLevel.LossCodesLevel1Id).Select(x => x.LossCode).FirstOrDefault();
                    reasonDetails.reasonLevel1 = reasonName;
                    reasonDetailsList.Add(reasonDetails);
                }
                else if (reasonLevel.LossCodesLevel == 2)
                {
                    //reasonName = reasonLevel.LossCode;
                    //reaons[0] = reasonName;
                    reasonDetails reasonDetails1 = new reasonDetails();
                    reasonDetails1.reasonLevel2 = reasonLevel.LossCode;
                    var reasonName = db.Tbllossescodes.Where(x => x.LossCodeId == reasonLevel.LossCodesLevel1Id).Select(x => x.LossCode).FirstOrDefault();
                    //reaons[1] = reasonName;
                    reasonDetails1.reasonLevel1 = reasonName;
                    reasonDetailsList.Add(reasonDetails1);
                }
                else if (reasonLevel.LossCodesLevel == 1)
                {
                    //reasonName = reasonLevel.LossCode;
                    //reaons[0] = reasonName;
                    reasonDetails reasonDetails = new reasonDetails();
                    reasonDetails.reasonLevel1 = reasonLevel.LossCode;
                    reasonDetailsList.Add(reasonDetails);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return reasonDetailsList;
        }

        /// <summary>
        /// Index
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 Index()
        {
            CommonResponse1 obj = new CommonResponse1();
            List<ModeIndex> listModeIndex = new List<ModeIndex>();
            try
            {
                var getTempModeData = db.TblTempMode.Where(x => x.IsDeleted == 0 && x.IsApproved != 0).ToList();
                if (getTempModeData.Count > 0)
                {
                    foreach (var row in getTempModeData)
                    {
                        ModeIndex objModeIndex = new ModeIndex();

                        if (row.IsApproved == 1)
                        {
                            objModeIndex.accpetReject = "";
                            objModeIndex.sentApproval = "Sent For Apporval";
                        }
                        else if (row.IsApproved == 2)
                        {
                            objModeIndex.accpetReject = "Accepted";
                            objModeIndex.sentApproval = "Sent For Apporval";
                        }
                        else if (row.IsApproved == 3)
                        {
                            objModeIndex.accpetReject = "Rejected";
                            objModeIndex.sentApproval = "Sent For Apporval";
                        }

                        objModeIndex.endDate = Convert.ToDateTime(row.EndTime).ToString("yyyy-MM-dd HH:mm:ss");
                        string tblMode = row.Mode;
                        objModeIndex.mode = tblMode;

                        if (tblMode == "IDLE" || tblMode == "BREAKDOWN")
                        {
                            int lossId = (int)db.TblTempLiveLossOfEntry.Where(x => x.TempModeId == row.TempModeId).Select(x => x.MessageCodeId).FirstOrDefault();
                            //string[] reasonsCode = GetReasonNames(lossId);
                            List<reasonDetails> reasonDetailsList = GetReasonNames(lossId);

                            foreach (var item in reasonDetailsList)
                            {
                                if (item.reasonLevel1 != null && item.reasonLevel2 == null && item.reasonLevel3 == null)
                                {
                                    objModeIndex.reasonLevel1 = item.reasonLevel1;
                                }
                                else if (item.reasonLevel1 != null && item.reasonLevel2 != null && item.reasonLevel3 == null)
                                {
                                    objModeIndex.reasonLevel1 = item.reasonLevel1;
                                    objModeIndex.reasonLevel2 = item.reasonLevel2;
                                }
                                else if (item.reasonLevel1 != null && item.reasonLevel2 != null && item.reasonLevel3 != null)
                                {
                                    objModeIndex.reasonLevel1 = item.reasonLevel1;
                                    objModeIndex.reasonLevel2 = item.reasonLevel2;
                                    objModeIndex.reasonLevel3 = item.reasonLevel3;
                                }
                            }
                        }
                        //    if (reasonsCode.Count() == 3)
                        //    {
                        //        for (int i = 0; i < reasonsCode.Count(); i++)
                        //        {
                        //            if (i == 0)
                        //            {
                        //                objModeIndex.reasonLevel1 = reasonsCode[i];
                        //            }
                        //            else if (i == 1)
                        //            {
                        //                objModeIndex.reasonLevel1 = reasonsCode[i];
                        //            }
                        //            else if (i == 2)
                        //            {
                        //                objModeIndex.reasonLevel1 = reasonsCode[i];
                        //            }
                        //        }
                        //    }
                        //    else if (reasonsCode.Count() == 2)
                        //    {
                        //        for (int i = 0; i < reasonsCode.Count(); i++)
                        //        {
                        //            if (i == 0)
                        //            {
                        //                objModeIndex.reasonLevel1 = reasonsCode[i];
                        //            }
                        //            else if (i == 1)
                        //            {
                        //                objModeIndex.reasonLevel1 = reasonsCode[i];
                        //            }

                        //        }
                        //    }
                        //    else if (reasonsCode.Count() == 1)
                        //    {
                        //        objModeIndex.reasonLevel1 = reasonsCode[0];
                        //    }


                        //}
                        //else
                        //{

                        //}

                        objModeIndex.startDate = Convert.ToDateTime(row.StartTime).ToString("yyyy-MM-dd HH:mm:ss");
                        objModeIndex.tempModeId = row.TempModeId;
                        listModeIndex.Add(objModeIndex);
                    }
                    obj.isStatus = true;
                    obj.response = listModeIndex;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";

                }
            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Add Individual Mode Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponseWithTempId AddIndividualModeDetails(AddIndividualMode data)
        {
            CommonResponseWithTempId obj = new CommonResponseWithTempId();
            try
            {
                #region For saving mode in temp mode before giving reason level
                if (data.lossCodeId == 0)
                {
                    var check = db.TblTempMode.Where(m => m.TempModeId == data.tempModeId && m.IsApproved == 0 && (m.OverAllSaved == null || m.OverAllSaved == 1)).FirstOrDefault();
                    if (check != null)
                    {
                        if (data.mode == "Idle")
                        {
                            check.ColorCode = "yellow";
                            check.Mode = "IDLE";
                        }
                        if (data.mode == "Breakdown")
                        {
                            check.ColorCode = "red";
                            check.Mode = "BREAKDOWN";
                        }
                        if (data.mode == "Running")
                        {
                            check.Mode = "Running";
                            check.ColorCode = "green";
                        }
                        if (data.mode == "PowerOff")
                        {
                            check.Mode = "PowerOff";
                            check.ColorCode = "blue";
                        }

                        check.Mode = check.Mode;
                        check.ColorCode = check.ColorCode;
                        db.SaveChanges();
                    }

                    if (check.Mode == "IDLE")
                    {

                        var level1Data = (from wf in db.Tbllossescodes
                                          where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType != "NoCode" && wf.MessageType != "BREAKDOWN" && wf.MessageType != "PM"
                                          select new
                                          {
                                              LossCodeId = wf.LossCodeId,
                                              LossCode = wf.LossCode
                                          }).ToList();

                        if (level1Data.Count != 0)
                        {
                            obj.isStatus = true;
                            obj.response = level1Data;
                            obj.tempModeId = data.tempModeId;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.response = "No Items Found";
                        }
                    }
                    else if (check.Mode == "BREAKDOWN")
                    {
                        var level1Data = (from wf in db.Tbllossescodes
                                          where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType == "BREAKDOWN" && wf.LossCode != "9999"
                                          select new
                                          {
                                              LossCode = wf.LossCode,
                                              LossCodeId = wf.LossCodeId,
                                          }).ToList();
                        if (level1Data.Count > 0)
                        {
                            obj.isStatus = true;
                            obj.response = level1Data;
                            obj.tempModeId = data.tempModeId;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.response = "No Items Found";
                        }
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "No Items Found";
                    }
                }
                #endregion

                #region Adding loss Code id in the temp livelossofEntry table 
                else
                {
                    var check = db.TblTempMode.Where(m => m.TempModeId == data.tempModeId && m.IsApproved == 0).FirstOrDefault();

                    DateTime Time = DateTime.Now;
                    TimeSpan Tm = new TimeSpan(Time.Hour, Time.Minute, Time.Second);
                    var ShiftDetails = db.TblshiftMstr.Where(m => m.StartTime <= Tm && m.EndTime >= Tm).FirstOrDefault();
                    string Shift = "C";
                    if (ShiftDetails != null)
                    {
                        Shift = ShiftDetails.ShiftName;
                    }

                    var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.lossCodeId).FirstOrDefault();

                    var tblTempLossOfEntry = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == check.TempModeId && m.CorrectedDate == check.CorrectedDate && m.MachineId == check.MachineId).FirstOrDefault();
                    if (tblTempLossOfEntry == null)
                    {
                        TblTempLiveLossOfEntry tblTempLiveLossOfEntry = new TblTempLiveLossOfEntry();
                        tblTempLiveLossOfEntry.StartDateTime = check.StartTime;
                        tblTempLiveLossOfEntry.EntryTime = check.StartTime;
                        tblTempLiveLossOfEntry.EndDateTime = check.EndTime;
                        tblTempLiveLossOfEntry.CorrectedDate = check.CorrectedDate;
                        tblTempLiveLossOfEntry.MachineId = check.MachineId;
                        tblTempLiveLossOfEntry.Shift = Shift;
                        tblTempLiveLossOfEntry.MessageCodeId = data.lossCodeId;
                        tblTempLiveLossOfEntry.MessageCode = lossdata.LossCode;
                        tblTempLiveLossOfEntry.MessageDesc = lossdata.LossCodeDesc;
                        tblTempLiveLossOfEntry.IsUpdate = 1;
                        tblTempLiveLossOfEntry.DoneWithRow = 1;
                        tblTempLiveLossOfEntry.IsStart = 0;
                        tblTempLiveLossOfEntry.IsScreen = 0;
                        tblTempLiveLossOfEntry.ForRefresh = 0;
                        tblTempLiveLossOfEntry.TempModeId = check.TempModeId;
                        tblTempLiveLossOfEntry.ModeId = check.ModeId;
                        db.TblTempLiveLossOfEntry.Add(tblTempLiveLossOfEntry);
                        db.SaveChanges();
                        obj.isStatus = true;
                        obj.response = "Added Successfully";
                    }
                    else
                    {
                        tblTempLossOfEntry.StartDateTime = check.StartTime;
                        tblTempLossOfEntry.EntryTime = check.StartTime;
                        tblTempLossOfEntry.EndDateTime = check.EndTime;
                        tblTempLossOfEntry.CorrectedDate = check.CorrectedDate;
                        tblTempLossOfEntry.MachineId = check.MachineId;
                        tblTempLossOfEntry.Shift = Shift;
                        tblTempLossOfEntry.MessageCodeId = data.lossCodeId;
                        tblTempLossOfEntry.MessageCode = lossdata.LossCode;
                        tblTempLossOfEntry.MessageDesc = lossdata.LossCodeDesc;
                        tblTempLossOfEntry.IsUpdate = 1;
                        tblTempLossOfEntry.DoneWithRow = 1;
                        tblTempLossOfEntry.IsStart = 0;
                        tblTempLossOfEntry.IsScreen = 0;
                        tblTempLossOfEntry.ForRefresh = 0;
                        tblTempLossOfEntry.TempModeId = check.TempModeId;
                        tblTempLossOfEntry.ModeId = check.ModeId;
                        db.SaveChanges();
                        obj.isStatus = true;
                        obj.response = "Updated Successfully";
                    }

                    var level2Data = (from wf in db.Tbllossescodes
                                      where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == data.lossCodeId && wf.LossCodesLevel == 2 && wf.LossCodesLevel2Id == null && wf.MessageType != "BREAKDOWN"
                                      select new
                                      {
                                          LossCodeId = wf.LossCodeId,
                                          LossCode = wf.LossCode,
                                      }).ToList();

                    var level3Data = (from wf in db.Tbllossescodes
                                      where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == data.lossCodeId && wf.LossCodesLevel2Id == null && wf.MessageType == "BREAKDOWN"
                                      select new
                                      {
                                          LossCode = wf.LossCode,
                                          LossCodeId = wf.LossCodeId,
                                      }).ToList();

                    if (level2Data.Count > 0)
                    {
                        obj.isStatus = true;
                        obj.response = level2Data;
                        obj.tempModeId = data.tempModeId;
                    }
                    else if (level3Data.Count > 0)
                    {
                        obj.isStatus = true;
                        obj.response = level3Data;
                        obj.tempModeId = data.tempModeId;
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "No Items Found";
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Get Reasons For Third Level
        /// </summary>
        /// <param name="tempModeId"></param>
        /// <param name="lossCodeId"></param>
        /// <returns></returns>
        public CommonResponseWithTempId GetReasonsForThirdLevel(int tempModeId, int lossCodeId)
        {
            CommonResponseWithTempId obj = new CommonResponseWithTempId();
            try
            {
                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == lossCodeId).FirstOrDefault();
                var tempData = (from wf in db.TblTempMode
                                where wf.TempModeId == tempModeId
                                select new
                                {
                                    CorrectedDate = wf.CorrectedDate,
                                    MachineId = wf.MachineId
                                }).FirstOrDefault();
                var tblTempLossOfEntry = db.TblTempLiveLossOfEntry.Where(m => m.TempModeId == tempModeId && m.CorrectedDate == tempData.CorrectedDate && m.MachineId == tempData.MachineId).FirstOrDefault();

                if (tblTempLossOfEntry != null)
                {
                    tblTempLossOfEntry.MessageCodeId = lossCodeId;
                    tblTempLossOfEntry.MessageCode = lossdata.LossCode;
                    tblTempLossOfEntry.MessageDesc = lossdata.LossCodeDesc;
                    db.SaveChanges();
                }

                var level3Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == lossCodeId && wf.LossCodesLevel == 3 && wf.MessageType != "BREAKDOWN"
                                  select new
                                  {
                                      LossCodeId = wf.LossCodeId,
                                      LossCode = wf.LossCode
                                  }).ToList();

                var level2Data = (from wf in db.Tbllossescodes
                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == lossCodeId && wf.MessageType == "BREAKDOWN"
                                  select new
                                  {
                                      LossCode = wf.LossCode,
                                      LossCodeId = wf.LossCodeId
                                  }).ToList();

                if (level3Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level3Data;
                    obj.tempModeId = tempModeId;
                }
                else if (level2Data.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = level2Data;
                    obj.tempModeId = tempModeId;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Get Time Details
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public CommonResponse1 GetTimeDetailsForSplitDuration(int tempModeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            List<ModeStartEndDateTime> modeStartEndDateTimesList = new List<ModeStartEndDateTime>();
            try
            {
                var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                if (check != null)
                {
                    DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
                    DateTime StartDateTime = Convert.ToDateTime(check.StartTime);

                    string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids = dt.Split();
                    string endDate = ids[0];
                    string endTime = ids[1];

                    string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids1 = dt1.Split();
                    string startDate = ids1[0];
                    string startTime = ids1[1];

                    #region Response assiging
                    ModeStartEndDateTime modeStartEndDateTime = new ModeStartEndDateTime();
                    modeStartEndDateTime.tempModeId = check.TempModeId;
                    modeStartEndDateTime.startDate = startDate;
                    modeStartEndDateTime.startTime = startTime;
                    modeStartEndDateTime.endDate = endDate;
                    modeStartEndDateTime.endTime = endTime;
                    modeStartEndDateTime.tempModeId = check.TempModeId;
                    modeStartEndDateTimesList.Add(modeStartEndDateTime);
                    obj.isStatus = true;
                    obj.response = modeStartEndDateTimesList;
                    #endregion
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Index Page
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 Index1()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                List<IndexDetails> indexDetailsList = new List<IndexDetails>();
                int machineId = 0;
                string machineName = "";
                string correctedDate = "";
                //var check = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved != 0).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).Select(m =>  m.ToList()).ToList();

                var check = (from tt in db.TblTempMode
                             where tt.IsDeleted == 0 && tt.IsApproved != 0
                             group tt by new { tt.CorrectedDate, tt.MachineId, tt.ApproveLevel } into ta
                             select new { item = ta.Key, count = ta.Count() }).ToList();


                foreach (var item in check)
                {
                    IndexDetails indexDetails = new IndexDetails();

                    machineId = item.item.MachineId;
                    correctedDate = item.item.CorrectedDate;
                    machineName = db.Tblmachinedetails.Where(m => m.MachineId == machineId).Select(m => m.MachineInvNo).FirstOrDefault();
                    var plantId = db.Tblmachinedetails.Where(m => m.MachineId == machineId).Select(m => m.PlantId).FirstOrDefault();
                    var cellId = db.Tblmachinedetails.Where(m => m.MachineId == machineId).Select(m => m.CellId).FirstOrDefault();
                    var shopId = db.Tblmachinedetails.Where(m => m.MachineId == machineId).Select(m => m.ShopId).FirstOrDefault();

                    indexDetails.plantName = db.Tblplant.Where(m => m.PlantId == plantId).Select(m => m.PlantName).FirstOrDefault();
                    indexDetails.shopName = db.Tblshop.Where(m => m.ShopId == shopId).Select(m => m.ShopName).FirstOrDefault();
                    indexDetails.cellName = db.Tblcell.Where(m => m.CellId == cellId).Select(m => m.CellName).FirstOrDefault();
                    indexDetails.machineName = machineName;
                    indexDetails.correctedDate = item.item.CorrectedDate;

                    var afterMailSent = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 1 && m.ApproveLevel == null && m.MachineId == machineId && m.CorrectedDate == correctedDate).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();

                    if (afterMailSent.Count > 0)
                    {
                        indexDetails.firstApproval = "Mail Sent and Approval is Pending";
                        indexDetailsList.Add(indexDetails);
                    }

                    var afterSecondMailSent = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 2 && m.ApproveLevel == 1 && m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved1 == 1).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();

                    if (afterSecondMailSent.Count > 0)
                    {
                        indexDetails.firstApproval = "Approved";
                        indexDetails.secondApproval = "Mail Sent and Approval is Pending";
                        indexDetailsList.Add(indexDetails);
                    }

                    var firstLevel = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 2 && m.ApproveLevel == 1 && m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved1 == null).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();

                    if (firstLevel.Count > 0)
                    {
                        indexDetails.firstApproval = "Approved";
                        indexDetails.secondApproval = "Mail Sent and Approval is Pending";
                        indexDetailsList.Add(indexDetails);
                    }

                    var secondLevel = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 2 && m.ApproveLevel == 2 && m.MachineId == machineId && m.CorrectedDate == correctedDate).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();
                    if (secondLevel.Count > 0)
                    {
                        indexDetails.firstApproval = "Approved";
                        indexDetails.secondApproval = "Approved";
                        indexDetailsList.Add(indexDetails);
                    }

                    var firstLevelReject = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 3 && m.ApproveLevel == null && m.MachineId == machineId && m.CorrectedDate == correctedDate).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();
                    if (firstLevelReject.Count > 0)
                    {
                        indexDetails.firstApproval = "Rejected";
                        indexDetailsList.Add(indexDetails);
                    }

                    var secondLevleReject = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.IsApproved == 3 && m.ApproveLevel == 1 && m.MachineId == machineId && m.CorrectedDate == correctedDate).Select(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).GroupBy(m => new { m.MachineId, m.ApproveLevel, m.CorrectedDate }).ToList();
                    if (secondLevleReject.Count > 0)
                    {
                        indexDetails.firstApproval = "Approved";
                        indexDetails.secondApproval = "Rejected";
                        indexDetailsList.Add(indexDetails);
                    }
                }
                obj.isStatus = true;
                obj.response = indexDetailsList;

            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// Split Duration
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public GeneralResponse1 SplitDuration(SplitDetails data)
        {
            GeneralResponse1 obj = new GeneralResponse1();
            try
            {
                List<ModeStartEndDateTime> modeStartEndDateTimeList = new List<ModeStartEndDateTime>();

                var check = db.TblTempMode.Where(m => m.TempModeId == data.tempModeId).FirstOrDefault();
                if (check != null)
                {
                    DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
                    DateTime StartDateTime = Convert.ToDateTime(check.StartTime);

                    //DateTime endT = Convert.ToDateTime(data.endDateTime);

                    string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids1 = dt.Split();
                    string endDate1 = ids1[0];
                    string endTime1 = ids1[1];

                    string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids2 = dt1.Split();
                    string startDate = ids2[0];
                    string startTime = ids2[1];

                    string endTimeLast = ids2[0] + " " + data.endTime;
                    DateTime endT = Convert.ToDateTime(endTimeLast);

                    #region jugad
                    EndDateTime = Convert.ToDateTime(check.EndTime);
                    StartDateTime = Convert.ToDateTime(check.StartTime);
                    bool flag = false;
                    if ((EndDateTime.Date - StartDateTime.Date).Days > 0)
                    {
                        flag = true;
                    }
                    #endregion

                    if (endT < EndDateTime && endT > StartDateTime)
                    {
                        check.EndTime = endT;
                        check.Mode = "PowerOff";
                        var durationInSec = (endT - StartDateTime).TotalSeconds;
                        check.DurationInSec = Convert.ToInt32(durationInSec);
                        db.SaveChanges();

                        #region add new row in temp mode table which we are going to insert in mode table
                        TblTempMode tblTempMode = new TblTempMode();
                        tblTempMode.MachineId = check.MachineId;
                        tblTempMode.ModeId = check.ModeId;
                        tblTempMode.Mode = "PowerOff";
                        tblTempMode.ColorCode = "blue";
                        tblTempMode.InsertedOn = DateTime.Now;
                        tblTempMode.InsertedBy = 1;
                        tblTempMode.IsSaved = 0;
                        tblTempMode.CorrectedDate = check.CorrectedDate;
                        tblTempMode.IsDeleted = 0;
                        tblTempMode.StartTime = endT;
                        tblTempMode.EndTime = EndDateTime;
                        tblTempMode.IsCompleted = 1;
                        tblTempMode.IsApproved = 0;
                        var durationInSec11 = (EndDateTime - endT).TotalSeconds;
                        tblTempMode.DurationInSec = Convert.ToInt32(durationInSec11);
                        db.TblTempMode.Add(tblTempMode);
                        db.SaveChanges();
                        #endregion

                        DateTime EndDateTime1 = Convert.ToDateTime(tblTempMode.EndTime);
                        DateTime StartDateTime1 = Convert.ToDateTime(tblTempMode.StartTime);

                        string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids3 = dt2.Split();
                        string endDate2 = ids3[0];
                        string endTime2 = ids3[1];

                        string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids4 = dt3.Split();
                        string startDate3 = ids4[0];
                        string startTime3 = ids4[1];

                        ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                        modeStartEndDateTime2.tempModeId = tblTempMode.TempModeId;
                        modeStartEndDateTime2.startDate = startDate3;
                        modeStartEndDateTime2.startTime = startTime3;
                        modeStartEndDateTime2.endDate = endDate2;
                        modeStartEndDateTime2.endTime = endTime2;
                        modeStartEndDateTime2.modeId = tblTempMode.ModeId;
                        modeStartEndDateTimeList.Add(modeStartEndDateTime2);
                    }
                    else if (flag == true)
                    {
                        string tempStartDate = StartDateTime.ToString("yyyy-MM-dd");
                        string tempEndDate = EndDateTime.ToString("yyyy-MM-dd");
                        DateTime endTT = Convert.ToDateTime(tempEndDate + " " + data.endTime);
                        string[] endTTString = EndDateTime.ToString().Split(' ');
                        DateTime endTT1 = Convert.ToDateTime(tempEndDate + " " + endTTString[1]);

                        if (endTT < EndDateTime && endTT > StartDateTime)
                        {
                            #region Update old row
                            check.EndTime = endTT;
                            check.Mode = "PowerOff";
                            var durationInSec = (endTT - StartDateTime).TotalSeconds;
                            check.DurationInSec = Convert.ToInt32(durationInSec);
                            db.SaveChanges();
                            #endregion

                            #region add new row in temp mode table which we are going to insert in mode table
                            TblTempMode tblTempMode = new TblTempMode();
                            tblTempMode.MachineId = check.MachineId;
                            tblTempMode.ModeId = check.ModeId;
                            tblTempMode.Mode = "PowerOff";
                            tblTempMode.ColorCode = "blue";
                            tblTempMode.InsertedOn = DateTime.Now;
                            tblTempMode.InsertedBy = 1;
                            tblTempMode.IsSaved = 0;
                            tblTempMode.CorrectedDate = check.CorrectedDate;
                            tblTempMode.IsDeleted = 0;
                            tblTempMode.StartTime = endTT;
                            tblTempMode.EndTime = EndDateTime;
                            tblTempMode.IsCompleted = 1;
                            tblTempMode.IsApproved = 0;
                            var durationInSec11 = (EndDateTime - endTT).TotalSeconds;
                            tblTempMode.DurationInSec = Convert.ToInt32(durationInSec11);
                            db.TblTempMode.Add(tblTempMode);
                            db.SaveChanges();
                            #endregion

                            DateTime EndDateTime1 = Convert.ToDateTime(tblTempMode.EndTime);
                            DateTime StartDateTime1 = Convert.ToDateTime(tblTempMode.StartTime);

                            string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                            string[] ids5 = dt2.Split();
                            string endDate2 = ids5[0];
                            string endTime2 = ids5[1];

                            string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                            string[] ids3 = dt3.Split();
                            string startDate3 = ids3[0];
                            string startTime3 = ids3[1];

                            ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                            modeStartEndDateTime2.tempModeId = tblTempMode.TempModeId;
                            modeStartEndDateTime2.startDate = startDate3;
                            modeStartEndDateTime2.startTime = startTime3;
                            modeStartEndDateTime2.endDate = endDate2;
                            modeStartEndDateTime2.endTime = endTime2;
                            modeStartEndDateTime2.modeId = tblTempMode.ModeId;
                            modeStartEndDateTimeList.Add(modeStartEndDateTime2);
                        }
                    }
                }

                string[] ids = data.tempModeIds.Split(',');
                foreach (var item in ids)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                    if (dbCheck != null)
                    {
                        DateTime EndDateTime1 = Convert.ToDateTime(dbCheck.EndTime);
                        DateTime StartDateTime1 = Convert.ToDateTime(dbCheck.StartTime);

                        string dt2 = EndDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids2 = dt2.Split();
                        string endDate2 = ids2[0];
                        string endTime2 = ids2[1];

                        string dt3 = StartDateTime1.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids3 = dt3.Split();
                        string startDate3 = ids3[0];
                        string startTime3 = ids3[1];

                        ModeStartEndDateTime modeStartEndDateTime2 = new ModeStartEndDateTime();
                        modeStartEndDateTime2.tempModeId = dbCheck.TempModeId;
                        modeStartEndDateTime2.startDate = startDate3;
                        modeStartEndDateTime2.startTime = startTime3;
                        modeStartEndDateTime2.endDate = endDate2;
                        modeStartEndDateTime2.endTime = endTime2;
                        modeStartEndDateTime2.modeId = dbCheck.ModeId;
                        modeStartEndDateTimeList.Add(modeStartEndDateTime2);
                    }
                }

                obj.isStatus = true;
                obj.response = modeStartEndDateTimeList.OrderBy(m => m.tempModeId);
                obj.tempModeIds = String.Join(",", modeStartEndDateTimeList.OrderBy(m => m.tempModeId).Select(m => m.tempModeId).ToList());
            }
            catch (Exception ex)
            {
                obj.isStatus = false;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="noLoginId"></param>
        /// <returns></returns>
        public GeneralResponse1 DeleteTempModeIds(DeleteTempMode data)
        {
            GeneralResponse1 obj = new GeneralResponse1();
            try
            {
                string[] ids = data.deleteTempModeIds.Split(',');
                int prevTempModeId = 0;
                for (int i = 0; i < ids.Length; i++)
                {
                    int tempModeId = Convert.ToInt32(ids[i]);
                    if (data.deleteTempModeId == tempModeId)
                    {
                        if ((i - 1) < 0)
                        {
                            prevTempModeId = Convert.ToInt32(ids[i + 1]);

                            var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();

                            var dbCheck = db.TblTempMode.Where(m => m.TempModeId == prevTempModeId).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                dbCheck.StartTime = check.StartTime;
                                dbCheck.EndTime = dbCheck.EndTime;
                                DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                var durationInSec = (EndDateTime - StartDateTime).TotalSeconds;
                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                                db.SaveChanges();

                                db.TblTempMode.Remove(check);
                                db.SaveChanges();
                            }
                        }
                        else
                        {
                            prevTempModeId = Convert.ToInt32(ids[i - 1]);

                            var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();

                            var dbCheck = db.TblTempMode.Where(m => m.TempModeId == prevTempModeId).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                dbCheck.EndTime = check.EndTime;
                                DateTime StartDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                var durationInSec = (EndDateTime - StartDateTime).TotalSeconds;
                                dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                                db.SaveChanges();

                                db.TblTempMode.Remove(check);
                                db.SaveChanges();
                            }
                        }
                    }
                }
                List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();

                foreach (var item in ids)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();

                    if (check != null)
                    {
                        DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
                        DateTime StartDateTime = Convert.ToDateTime(check.StartTime);

                        string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids1 = dt.Split();
                        string endDate = ids1[0];
                        string endTime = ids1[1];

                        string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids2 = dt1.Split();
                        string startDate = ids2[0];
                        string startTime = ids2[1];

                        ModeStartEndDateTime noLoginStartEndDateTime = new ModeStartEndDateTime();
                        noLoginStartEndDateTime.startDate = startDate;
                        noLoginStartEndDateTime.startTime = startTime;
                        noLoginStartEndDateTime.endDate = endDate;
                        noLoginStartEndDateTime.endTime = endTime;
                        noLoginStartEndDateTime.tempModeId = check.TempModeId;
                        noLoginStartEndDateTime.modeId = check.ModeId;
                        modeStartEndDateTimes.Add(noLoginStartEndDateTime);
                    }
                }
                obj.isStatus = true;
                obj.response = modeStartEndDateTimes;
                obj.tempModeIds = String.Join(",", modeStartEndDateTimes.OrderBy(m => m.tempModeId).Select(m => m.tempModeId).ToList());
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Update In BackUp Tables
        /// </summary>
        /// <param name="machineId"></param>
        /// <param name="correctedDate"></param>
        /// <returns></returns>
        public bool UpdateInBackUpTables(int machineId, string correctedDate)
        {
            bool flag = false;

            var check = (from wf in db.Tblhmiscreen
                         where wf.MachineId == machineId && wf.CorrectedDate == correctedDate
                         select new
                         {
                             Hmiid = wf.Hmiid
                         }).ToList();

            foreach (var item in check)
            {
                int hmiid = Convert.ToInt32(item.Hmiid);
                var dbCheck = db.Tblwolossess.Where(m => m.Hmiid == hmiid).FirstOrDefault();
                if (dbCheck != null)
                {
                    TblwolossessBackup tblwolossessBackup = new TblwolossessBackup();
                    tblwolossessBackup.WoLossesId = dbCheck.WolossesId;
                    tblwolossessBackup.Hmiid = dbCheck.Hmiid;
                    tblwolossessBackup.LossId = dbCheck.LossId;
                    tblwolossessBackup.LossName = dbCheck.LossName;
                    tblwolossessBackup.LossDuration = dbCheck.LossDuration;
                    tblwolossessBackup.Level = dbCheck.Level;
                    tblwolossessBackup.LossCodeLevel1Id = dbCheck.LossCodeLevel1Id;
                    tblwolossessBackup.LossCodeLevel1Name = dbCheck.LossCodeLevel1Name;
                    tblwolossessBackup.LossCodeLevel2Id = dbCheck.LossCodeLevel2Id;
                    tblwolossessBackup.LossCodeLevel2Name = dbCheck.LossCodeLevel2Name;
                    tblwolossessBackup.InsertedOn = DateTime.Now;
                    tblwolossessBackup.IsDeleted = 0;
                    db.TblwolossessBackup.Add(tblwolossessBackup);
                    db.SaveChanges();
                    db.Tblwolossess.Remove(dbCheck);
                    db.SaveChanges();
                }
            }

            var check1 = db.Tblworeport.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate).ToList();

            foreach (var item1 in check1)
            {
                TblworeportBackup tblworeportBackup = new TblworeportBackup();
                tblworeportBackup.WoReportId = item1.WoreportId;
                tblworeportBackup.MachineId = item1.MachineId;
                tblworeportBackup.Hmiid = item1.Hmiid;
                tblworeportBackup.OperatorName = item1.OperatorName;
                tblworeportBackup.Shift = item1.Shift;
                tblworeportBackup.CorrectedDate = item1.CorrectedDate;
                tblworeportBackup.PartNo = item1.PartNo;
                tblworeportBackup.WorkOrderNo = item1.WorkOrderNo;
                tblworeportBackup.OpNo = item1.OpNo;
                tblworeportBackup.TargetQty = item1.TargetQty;
                tblworeportBackup.DeliveredQty = item1.DeliveredQty;
                tblworeportBackup.IsPf = item1.IsPf;
                tblworeportBackup.IsHold = item1.IsHold;
                tblworeportBackup.CuttingTime = item1.CuttingTime;
                tblworeportBackup.SettingTime = item1.SettingTime;
                tblworeportBackup.SelfInspection = item1.SelfInspection;
                tblworeportBackup.Idle = item1.Idle;
                tblworeportBackup.Breakdown = item1.Breakdown;
                tblworeportBackup.Type = item1.Type;
                tblworeportBackup.NccuttingTimePerPart = item1.NccuttingTimePerPart;
                tblworeportBackup.TotalNccuttingTime = item1.TotalNccuttingTime;
                tblworeportBackup.Woefficiency = item1.Woefficiency;
                tblworeportBackup.RejectedQty = item1.RejectedQty;
                tblworeportBackup.RejectedReason = item1.RejectedReason;
                tblworeportBackup.Program = item1.Program;
                tblworeportBackup.Mrweight = item1.Mrweight;
                tblworeportBackup.InsertedOn = DateTime.Now;
                tblworeportBackup.IsMultiWo = item1.IsMultiWo;
                tblworeportBackup.IsNormalWc = item1.IsNormalWc;
                tblworeportBackup.HoldReason = item1.HoldReason;
                tblworeportBackup.MinorLoss = item1.MinorLoss;
                tblworeportBackup.SplitWo = item1.SplitWo;
                tblworeportBackup.Blue = item1.Blue;
                tblworeportBackup.ScrapQtyTime = item1.ScrapQtyTime;
                tblworeportBackup.ReWorkTime = item1.ReWorkTime;
                tblworeportBackup.SummationOfSctvsPp = item1.SummationOfSctvsPp;
                tblworeportBackup.StartTime = item1.StartTime;
                tblworeportBackup.EndTime = item1.EndTime;
                tblworeportBackup.BatchNo = item1.BatchNo;
                db.TblworeportBackup.Add(tblworeportBackup);
                db.SaveChanges();
            }
            db.Tblworeport.RemoveRange(check1);
            db.SaveChanges();


            DateTime corrDate = Convert.ToDateTime(correctedDate + " " + "00:00:00");
            var check2 = db.Tbloeedashboardvariables.Where(m => m.Wcid == machineId && m.StartDate == corrDate).FirstOrDefault();
            if (check2 != null)
            {
                TbloeeDashBoardVbackUp tbloeeDashBoardVbackUp = new TbloeeDashBoardVbackUp();
                tbloeeDashBoardVbackUp.OeeVariablesId = check2.OeevariablesId;
                tbloeeDashBoardVbackUp.PlantId = check2.PlantId;
                tbloeeDashBoardVbackUp.ShopId = check2.ShopId;
                tbloeeDashBoardVbackUp.CellId = check2.CellId;
                tbloeeDashBoardVbackUp.Wcid = check2.Wcid;
                tbloeeDashBoardVbackUp.StartDate = check2.StartDate;
                tbloeeDashBoardVbackUp.EndDate = check2.EndDate;
                tbloeeDashBoardVbackUp.MinorLosses = check2.MinorLosses;
                tbloeeDashBoardVbackUp.Blue = check2.Blue;
                tbloeeDashBoardVbackUp.Green = check2.Green;
                tbloeeDashBoardVbackUp.SettingTime = check2.SettingTime;
                tbloeeDashBoardVbackUp.RoaLossess = check2.Roalossess;
                tbloeeDashBoardVbackUp.DownTimeBreakdown = check2.DownTimeBreakdown;
                tbloeeDashBoardVbackUp.SummationOfSctvsPp = check2.SummationOfSctvsPp;
                tbloeeDashBoardVbackUp.ScrapQtyTime = check2.ScrapQtyTime;
                tbloeeDashBoardVbackUp.ReWotime = check2.ReWotime;
                tbloeeDashBoardVbackUp.Loss1Name = check2.Loss1Name;
                tbloeeDashBoardVbackUp.Loss1Value = check2.Loss1Value;
                tbloeeDashBoardVbackUp.Loss2Name = check2.Loss2Name;
                tbloeeDashBoardVbackUp.Loss2Value = check2.Loss2Value;
                tbloeeDashBoardVbackUp.Loss3Name = check2.Loss3Name;
                tbloeeDashBoardVbackUp.Loss3Value = check2.Loss3Value;
                tbloeeDashBoardVbackUp.Loss4Name = check2.Loss4Name;
                tbloeeDashBoardVbackUp.Loss4Value = check2.Loss4Value;
                tbloeeDashBoardVbackUp.Loss5Name = check2.Loss5Name;
                tbloeeDashBoardVbackUp.Loss5Value = check2.Loss5Value;
                tbloeeDashBoardVbackUp.CreatedOn = DateTime.Now;
                tbloeeDashBoardVbackUp.CreatedBy = 1;
                tbloeeDashBoardVbackUp.IsDeleted = check2.IsDeleted;
                db.TbloeeDashBoardVbackUp.Add(tbloeeDashBoardVbackUp);
                db.SaveChanges();
                flag = true;
                db.Tbloeedashboardvariables.Remove(check2);
                db.SaveChanges();
            }
            return flag;
        }

        /// <summary>
        /// Update In Report
        /// </summary>
        /// <param name="machineId"></param>
        /// <param name="correctedDate"></param>
        /// <returns></returns>
        //public CommonResponse1 UpdateInReport(int machineId, string correctedDate)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        bool result = false;
        //        result = UpdateInBackUpTables(machineId, correctedDate);
        //        if (result == true)
        //        {
        //            DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
        //            DateTime correcteDateTime = Convert.ToDateTime(correctedDate);
        //            Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(correcteDateTime, correcteDateTime);  // for WO report updation
        //            Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(correcteDateTime, correcteDateTime);// for OEE report updation
        //            var check = db.TblTempMode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved == 2 && m.ApproveLevel == 1).ToList();
        //            foreach (var item in check)
        //            {
        //                int tempModeId = Convert.ToInt32(item.TempModeId);
        //                var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
        //                if (dbCheck != null)
        //                {
        //                    dbCheck.IsUpdated = 1;
        //                    db.SaveChanges();
        //                }
        //            }

        //            var check1 = db.TblTempMode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved == 2 && m.ApproveLevel == 2).ToList();
        //            foreach (var item in check1)
        //            {
        //                int tempModeId = Convert.ToInt32(item.TempModeId);
        //                var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
        //                if (dbCheck != null)
        //                {
        //                    dbCheck.IsUpdated = 1;
        //                    db.SaveChanges();
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isStatus = false;
        //    }
        //    return obj;
        //}

        #region New UpdateInReport

        /// <summary>
        /// Update In Report
        /// </summary>
        /// <param name="machineId"></param>
        /// <param name="correctedDate"></param>
        /// <returns></returns>
        public CommonResponse1 UpdateInReport(int machineId, string correctedDate)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                bool result = false;
                result = UpdateInBackUpTables(machineId, correctedDate);

                DALCommonMethod commonMethodObj = new DALCommonMethod(db, configuration);
                DateTime correcteDateTime = Convert.ToDateTime(correctedDate);
                List<int> machinelist = new List<int>();
                machinelist.Add(machineId);
                Task<bool> reportWOUpdate = commonMethodObj.CalWODataForYesterday(correcteDateTime, correcteDateTime, machinelist);  // for WO report updation
                Task<bool> reportOEEUpdate = commonMethodObj.CalculateOEEForYesterday(correcteDateTime, correcteDateTime, machinelist);// for OEE report updation
                var check = db.TblTempMode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved == 2 && m.ApproveLevel == 1).ToList();
                foreach (var item in check)
                {
                    int tempModeId = Convert.ToInt32(item.TempModeId);
                    var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                    if (dbCheck != null)
                    {
                        dbCheck.IsUpdated = 1;
                        db.SaveChanges();
                    }
                }

                var check1 = db.TblTempMode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.IsApproved == 2 && m.ApproveLevel == 2).ToList();
                foreach (var item in check1)
                {
                    int tempModeId = Convert.ToInt32(item.TempModeId);
                    var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                    if (dbCheck != null)
                    {
                        dbCheck.IsUpdated = 1;
                        db.SaveChanges();
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

        #endregion

        /// <summary>
        /// Update Mimics Details
        /// </summary>
        /// <param name="machineId"></param>
        /// <param name="correctedDate"></param>
        /// <returns></returns>
        public CommonResponse1 UpdateMimicsDetails(int machineId, string correctedDate)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                bool result = false;
                bool flag = false;
                result = UpdateInMimicsBackUpTables(machineId, correctedDate);
                flag = UpdateMimicsdet(machineId, correctedDate);
                if (flag == true)
                {
                    obj.isStatus = true;
                    obj.response = "Updated Successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        public bool UpdateInMimicsBackUpTables(int machineId, string correctedDate)
        {
            bool flag = false;
            try
            {
                var check = db.Tblmimics.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate).ToList();
                if (check.Count > 0)
                {
                    foreach (var item in check)
                    {
                        TblMimicsBackup tblMimicsBackup = new TblMimicsBackup();
                        tblMimicsBackup.Mid = item.Mid;
                        tblMimicsBackup.MachineId = item.MachineId;
                        tblMimicsBackup.MachineOnTime = item.MachineOnTime;
                        tblMimicsBackup.OperatingTime = item.OperatingTime;
                        tblMimicsBackup.SetupTime = item.SetupTime;
                        tblMimicsBackup.IdleTime = item.IdleTime;
                        tblMimicsBackup.MachineOffTime = item.MachineOffTime;
                        tblMimicsBackup.BreakdownTime = item.BreakdownTime;
                        tblMimicsBackup.Shift = item.Shift;
                        tblMimicsBackup.CorrectedDate = item.CorrectedDate;
                        db.TblMimicsBackup.Add(tblMimicsBackup);
                        db.SaveChanges();
                    }
                    db.Tblmimics.RemoveRange(check);
                    db.SaveChanges();
                    flag = true;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return flag;
        }

        private bool UpdateMimicsdet(int machineId, string correctedDate)
        {
            bool flag = false;
            decimal OperatingTime = 0;
            decimal LossTime = 0;
            decimal MntTime = 0;
            decimal SetupTime = 0;
            decimal SetupMinorTime = 0;
            decimal PowerOffTime = 0;
            decimal PowerONTime = 0;
            //int MachineID = 24;
            var GetModeDurations = new List<Tblmode>();
            DateTime corrDate = Convert.ToDateTime(correctedDate);
            string correctedDatestr = corrDate.ToString("yyyy-MM-dd");

            using (i_facility_talContext db = new i_facility_talContext())
            {
                GetModeDurations = db.Tblmode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDatestr && m.IsCompleted == 1).ToList();
            }

            OperatingTime = Convert.ToDecimal(GetModeDurations.Where(m => m.ColorCode == "green").ToList().Sum(m => m.DurationInSec));
            PowerOffTime = Convert.ToDecimal(GetModeDurations.Where(m => m.ColorCode == "blue").ToList().Sum(m => m.DurationInSec));
            MntTime = Convert.ToDecimal(GetModeDurations.Where(m => m.ColorCode == "red").ToList().Sum(m => m.DurationInSec));
            LossTime = Convert.ToDecimal(GetModeDurations.Where(m => m.ColorCode == "yellow").ToList().Sum(m => m.DurationInSec));


            var GetModeDurationsRunning = new List<Tblmode>();
            using (i_facility_talContext db = new i_facility_talContext())
            {
                GetModeDurationsRunning = db.Tblmode.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDatestr && m.IsCompleted == 0).ToList();
            }

            foreach (var ModeRow in GetModeDurationsRunning)
            {
                String ColorCode = ModeRow.ColorCode;
                DateTime StartTime = (DateTime)ModeRow.StartTime;
                decimal Duration = (decimal)System.DateTime.Now.Subtract(StartTime).TotalSeconds;

                if (ColorCode == "yellow" || ColorCode == "YELLOW")
                {
                    LossTime += Duration;
                }
                else if (ColorCode == "green" || ColorCode == "GREEN")
                {
                    OperatingTime += Duration;
                }
                else if (ColorCode == "red" || ColorCode == "RED")
                {
                    MntTime += Duration;
                }
                else if (ColorCode == "blue" || ColorCode == "BLUE")
                {
                    PowerOffTime += Duration;
                }
            }

            PowerONTime = OperatingTime + LossTime;
            int IdleTime = Convert.ToInt32(LossTime + SetupTime + SetupMinorTime);
            OperatingTime = Math.Round((OperatingTime / 60), 2);
            PowerOffTime = (PowerOffTime / 60);
            MntTime = (MntTime / 60);
            IdleTime = (IdleTime / 60);
            PowerONTime = (PowerONTime / 60);


            //string correctedDt = correctedDate.Date.ToString("yyyy-MM-dd");
            var mimicsdata = new Tblmimics();
            using (i_facility_talContext db = new i_facility_talContext())
            {
                mimicsdata = db.Tblmimics.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDatestr).FirstOrDefault();
            }
            if (mimicsdata == null)
            {
                Tblmimics mimics = new Tblmimics();
                mimics.MachineId = machineId;
                mimics.CorrectedDate = correctedDatestr;
                mimics.OperatingTime = Convert.ToInt32(OperatingTime).ToString();
                mimics.BreakdownTime = Convert.ToInt32(MntTime).ToString();
                mimics.MachineOffTime = Convert.ToInt32(PowerOffTime).ToString();
                mimics.IdleTime = Convert.ToInt32(IdleTime).ToString();
                mimics.SetupTime = Convert.ToInt32(SetupTime).ToString();
                mimics.MachineOnTime = Convert.ToInt32(PowerONTime).ToString();
                using (i_facility_talContext db = new i_facility_talContext())
                {
                    db.Tblmimics.Add(mimics);
                    db.SaveChanges();
                    flag = true;
                }

            }
            else
            {
                var mimicsrow = new Tblmimics();
                using (i_facility_talContext db = new i_facility_talContext())
                {
                    mimicsrow = db.Tblmimics.Find(mimicsdata.Mid);
                }
                mimicsrow.MachineId = machineId;
                mimicsrow.CorrectedDate = correctedDatestr;
                mimicsrow.OperatingTime = Convert.ToInt32(OperatingTime).ToString();
                mimicsrow.BreakdownTime = Convert.ToInt32(MntTime).ToString().ToString();
                mimicsrow.MachineOffTime = Convert.ToInt32(PowerOffTime).ToString();
                mimicsrow.IdleTime = Convert.ToInt32(IdleTime).ToString();
                mimicsrow.SetupTime = Convert.ToInt32(SetupTime).ToString();
                mimicsrow.MachineOnTime = Convert.ToInt32(PowerONTime).ToString();
                db.SaveChanges();
                flag = true;
                //using (i_facility_talContext db = new i_facility_talContext())
                //{
                //    db.Entry(mimicsrow).State = System.Data.Entity.EntityState.Modified;
                //    db.SaveChanges();
                //}
            }
            return flag;
        }

        /// <summary>
        /// Add Reasons for mode Details
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 AddLossDetails(int tempModeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();
                if (check != null)
                {
                    check.OverAllSaved = 1;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Saved Successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Update Split Duration
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public GeneralResponse1 UpdateSplitDuration(UpdateTempMode data)
        {
            GeneralResponse1 obj = new GeneralResponse1();
            try
            {
                var check = db.TblTempMode.Where(m => m.TempModeId == data.updateTempModeId).FirstOrDefault();
                if (check != null)
                {
                    DateTime EndDateTime = Convert.ToDateTime(check.EndTime);
                    DateTime StartDateTime = Convert.ToDateTime(check.StartTime);

                    //DateTime endT = Convert.ToDateTime(data.endDateTime);

                    string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids1 = dt.Split();
                    string endDate1 = ids1[0];
                    string endTime1 = ids1[1];

                    string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    string[] ids2 = dt1.Split();
                    string startDate = ids2[0];
                    string startTime = ids2[1];

                    string endTimeLast = ids1[0] + " " + data.endTime;
                    DateTime endT = Convert.ToDateTime(endTimeLast);

                    check.EndTime = endT;
                    var durationInSec = (endT - StartDateTime).TotalSeconds;
                    check.DurationInSec = Convert.ToInt32(durationInSec);
                    db.SaveChanges();
                }

                string[] ids = data.updateTempModeIds.Split(',');

                int AfterTempModeId = 0;
                for (int i = 0; i < ids.Length; i++)
                {
                    int tempModeId = Convert.ToInt32(ids[i]);
                    if (data.updateTempModeId == tempModeId)
                    {
                        AfterTempModeId = Convert.ToInt32(ids[i + 1]);
                        var dbCheck = db.TblTempMode.Where(m => m.TempModeId == AfterTempModeId).FirstOrDefault();
                        if (dbCheck != null)
                        {
                            DateTime EndDateTime = Convert.ToDateTime(dbCheck.EndTime);
                            DateTime StartDateTime = Convert.ToDateTime(check.EndTime);

                            dbCheck.StartTime = StartDateTime;
                            var durationInSec = (EndDateTime - StartDateTime).TotalSeconds;
                            dbCheck.DurationInSec = Convert.ToInt32(durationInSec);
                            db.SaveChanges();
                        }
                    }
                }

                List<ModeStartEndDateTime> modeStartEndDateTimes = new List<ModeStartEndDateTime>();

                foreach (var item in ids)
                {
                    int tempModeId = Convert.ToInt32(item);
                    var checks = db.TblTempMode.Where(m => m.TempModeId == tempModeId).FirstOrDefault();

                    if (checks != null)
                    {
                        DateTime EndDateTime = Convert.ToDateTime(checks.EndTime);
                        DateTime StartDateTime = Convert.ToDateTime(checks.StartTime);

                        string dt = EndDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids1 = dt.Split();
                        string endDate = ids1[0];
                        string endTime = ids1[1];

                        string dt1 = StartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                        string[] ids2 = dt1.Split();
                        string startDate = ids2[0];
                        string startTime = ids2[1];

                        ModeStartEndDateTime noLoginStartEndDateTime = new ModeStartEndDateTime();
                        noLoginStartEndDateTime.startDate = startDate;
                        noLoginStartEndDateTime.startTime = startTime;
                        noLoginStartEndDateTime.endDate = endDate;
                        noLoginStartEndDateTime.endTime = endTime;
                        noLoginStartEndDateTime.tempModeId = checks.TempModeId;
                        noLoginStartEndDateTime.modeId = checks.ModeId;
                        modeStartEndDateTimes.Add(noLoginStartEndDateTime);
                    }
                }
                obj.isStatus = true;
                obj.response = modeStartEndDateTimes;
                obj.tempModeIds = String.Join(",", modeStartEndDateTimes.OrderBy(m => m.tempModeId).Select(m => m.tempModeId).ToList());
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Update Losses
        /// </summary>
        /// <param name="tempodeId"></param>
        /// <returns></returns>
        public CommonResponse1 UpdateLosses(int tempodeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblTempMode.Where(m => m.TempModeId == tempodeId).FirstOrDefault();
                if (check != null)
                {
                    check.IsUpdateFinal = 1;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Updated Successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Approved Info
        /// </summary>
        /// <param name="tempModeIds"></param>
        /// <returns></returns>
        public CommonResponse1 GetApprovedDetails(SendMailDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            List<ApproveReject> approveRejectsList = new List<ApproveReject>();
            try
            {
                string[] ids = data.tempModeIds.Split(',');
                if (data.checkIds == "")
                {
                    var check = db.TblTempMode.Where(m => m.IsDeleted == 0 && m.CorrectedDate == data.date && m.MachineId == data.machineId && m.IsHold == 1).ToList();
                    if (check.Count > 0)
                    {
                        foreach (var item in check)
                        {
                            var lossData = (from wf in db.TblTempLiveLossOfEntry
                                            join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                            where wf.TempModeId == item.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == item.CorrectedDate
                                            select new
                                            {
                                                LossCodeId = lc.LossCodeId,
                                                LossCode = lc.LossCode,
                                                LossCodesLevel = lc.LossCodesLevel,
                                                LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                StartDateTime = wf.StartDateTime,
                                                EndDateTime = wf.EndDateTime,
                                                MachineId = wf.MachineId,
                                                MessageCodeId = wf.MessageCodeId,
                                                CorrectedDate = wf.CorrectedDate,
                                                TempLossId = wf.TempLossId,

                                                ress1 = (lc.LossCodesLevel == 1 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                ress2 = (lc.LossCodesLevel == 2 ?
                                                 lc.LossCode :
                                                 db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                            }).ToList();

                            if (lossData.Count > 0)
                            {
                                foreach (var items in lossData)
                                {
                                    ApproveReject approveReject = new ApproveReject();
                                    approveReject.reason1 = items.ress1;
                                    approveReject.reason2 = items.ress2;
                                    approveReject.reason3 = items.ress3;
                                    approveReject.machineId = item.MachineId;
                                    approveReject.mode = item.Mode;
                                    approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == item.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                    DateTime startDateTime = Convert.ToDateTime(item.StartTime);
                                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.startTime = dt;
                                    DateTime endDateTime = Convert.ToDateTime(item.EndTime);
                                    string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.endTime = dt1;
                                    approveReject.tempModeId = item.TempModeId;
                                    approveRejectsList.Add(approveReject);
                                }
                            }
                            else
                            {
                                ApproveReject approveReject = new ApproveReject();
                                approveReject.machineId = item.MachineId;
                                approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == item.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                approveReject.mode = item.Mode;
                                //approveReject.startTime = check.StartTime.ToString();
                                //approveReject.endTime = check.EndTime.ToString();
                                DateTime startDateTime = Convert.ToDateTime(item.StartTime);
                                string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                approveReject.startTime = dt;
                                DateTime endDateTime = Convert.ToDateTime(item.EndTime);
                                string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                approveReject.endTime = dt1;
                                approveReject.tempModeId = item.TempModeId;
                                approveRejectsList.Add(approveReject);
                            }

                        }
                    }
                    else
                    {
                        obj.isStatus = false;
                        obj.response = "All The Reasons are Accepted";
                    }
                }
                else
                {
                    bool ret = false;
                    foreach (var irem1 in ids)
                    {
                        int tempModeId = Convert.ToInt32(irem1);
                        var dbCheck = db.TblTempMode.Where(x => x.TempModeId == tempModeId && x.IsHold == 0).FirstOrDefault();
                        if (dbCheck != null)
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
                        foreach(var items2 in ids)
                        {
                            int tempModeId = Convert.ToInt32(items2);
                            var dbCheck = db.TblTempMode.Where(m => m.TempModeId == tempModeId && m.IsHold == 0 && (m.ApproveLevel == 1 || m.ApproveLevel == 2)).FirstOrDefault();
                            if(dbCheck != null)
                            {
                                var lossData = (from wf in db.TblTempLiveLossOfEntry
                                                join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                                where wf.TempModeId == dbCheck.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == dbCheck.CorrectedDate
                                                select new
                                                {
                                                    LossCodeId = lc.LossCodeId,
                                                    LossCode = lc.LossCode,
                                                    LossCodesLevel = lc.LossCodesLevel,
                                                    LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                    LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                    StartDateTime = wf.StartDateTime,
                                                    EndDateTime = wf.EndDateTime,
                                                    MachineId = wf.MachineId,
                                                    MessageCodeId = wf.MessageCodeId,
                                                    CorrectedDate = wf.CorrectedDate,
                                                    TempLossId = wf.TempLossId,

                                                    ress1 = (lc.LossCodesLevel == 1 ?
                                                     lc.LossCode :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                    ress2 = (lc.LossCodesLevel == 2 ?
                                                     lc.LossCode :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                    ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                                }).ToList();

                                if (lossData.Count > 0)
                                {
                                    foreach (var items in lossData)
                                    {
                                        ApproveReject approveReject = new ApproveReject();
                                        approveReject.reason1 = items.ress1;
                                        approveReject.reason2 = items.ress2;
                                        approveReject.reason3 = items.ress3;
                                        approveReject.machineId = dbCheck.MachineId;
                                        approveReject.mode = dbCheck.Mode;
                                        approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId ==  dbCheck.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                        DateTime startDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                        approveReject.startTime = dt;
                                        DateTime endDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                        string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                        approveReject.endTime = dt1;
                                        approveReject.tempModeId = dbCheck.TempModeId;
                                        approveRejectsList.Add(approveReject);
                                    }
                                }
                                else
                                {
                                    ApproveReject approveReject = new ApproveReject();
                                    approveReject.machineId = dbCheck.MachineId;
                                    approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                    approveReject.mode = dbCheck.Mode;
                                    //approveReject.startTime = check.StartTime.ToString();
                                    //approveReject.endTime = check.EndTime.ToString();
                                    DateTime startDateTime = Convert.ToDateTime(dbCheck.StartTime);
                                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.startTime = dt;
                                    DateTime endDateTime = Convert.ToDateTime(dbCheck.EndTime);
                                    string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.endTime = dt1;
                                    approveReject.tempModeId = dbCheck.TempModeId;
                                    approveRejectsList.Add(approveReject);
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var items3 in ids)
                        {
                            int tempModeId = Convert.ToInt32(items3);
                            var dbCheck1 = db.TblTempMode.Where(x => x.TempModeId == tempModeId && x.IsHold == 1).FirstOrDefault();
                            if(dbCheck1 != null)
                            {
                                var lossData = (from wf in db.TblTempLiveLossOfEntry
                                                join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                                                where wf.TempModeId == dbCheck1.TempModeId && wf.DoneWithRow == 1 && wf.MachineId == data.machineId && wf.CorrectedDate == dbCheck1.CorrectedDate
                                                select new
                                                {
                                                    LossCodeId = lc.LossCodeId,
                                                    LossCode = lc.LossCode,
                                                    LossCodesLevel = lc.LossCodesLevel,
                                                    LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                                    LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                                    StartDateTime = wf.StartDateTime,
                                                    EndDateTime = wf.EndDateTime,
                                                    MachineId = wf.MachineId,
                                                    MessageCodeId = wf.MessageCodeId,
                                                    CorrectedDate = wf.CorrectedDate,
                                                    TempLossId = wf.TempLossId,

                                                    ress1 = (lc.LossCodesLevel == 1 ?
                                                     lc.LossCode :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                                    ress2 = (lc.LossCodesLevel == 2 ?
                                                     lc.LossCode :
                                                     db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                                    ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),
                                                }).ToList();

                                if (lossData.Count > 0)
                                {
                                    foreach (var items in lossData)
                                    {
                                        ApproveReject approveReject = new ApproveReject();
                                        approveReject.reason1 = items.ress1;
                                        approveReject.reason2 = items.ress2;
                                        approveReject.reason3 = items.ress3;
                                        approveReject.machineId = dbCheck1.MachineId;
                                        approveReject.mode = dbCheck1.Mode;
                                        approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck1.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                        DateTime startDateTime = Convert.ToDateTime(dbCheck1.StartTime);
                                        string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                        approveReject.startTime = dt;
                                        DateTime endDateTime = Convert.ToDateTime(dbCheck1.EndTime);
                                        string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                        approveReject.endTime = dt1;
                                        approveReject.tempModeId = dbCheck1.TempModeId;
                                        approveRejectsList.Add(approveReject);
                                    }
                                }
                                else
                                {
                                    ApproveReject approveReject = new ApproveReject();
                                    approveReject.machineId = dbCheck1.MachineId;
                                    approveReject.machineName = db.Tblmachinedetails.Where(m => m.MachineId == dbCheck1.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                    approveReject.mode = dbCheck1.Mode;
                                    //approveReject.startTime = check.StartTime.ToString();
                                    //approveReject.endTime = check.EndTime.ToString();
                                    DateTime startDateTime = Convert.ToDateTime(dbCheck1.StartTime);
                                    string dt = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.startTime = dt;
                                    DateTime endDateTime = Convert.ToDateTime(dbCheck1.EndTime);
                                    string dt1 = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                                    approveReject.endTime = dt1;
                                    approveReject.tempModeId = dbCheck1.TempModeId;
                                    approveRejectsList.Add(approveReject);
                                }
                            }
                        }
                    }
                }


                obj.isStatus = true;
                obj.response = approveRejectsList;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

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
                            var check1 = db.TblTempMode.Where(m => intArry.Contains(m.TempModeId) && m.IsApproved == 1).Select(m => new { m.IsApproved, m.ApproveLevel }).GroupBy(m => new { m.IsApproved, m.ApproveLevel }).ToList();
                            if (check1.Count > 0)
                            {
                                var toMail1 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 3).Select(m => m.FirstApproverToList).FirstOrDefault();

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

                            var check2 = db.TblTempMode.Where(m => intArry.Contains(m.TempModeId) && m.ApproveLevel == 1).Select(m => new { m.IsApproved, m.ApproveLevel }).GroupBy(m => new { m.IsApproved, m.ApproveLevel }).ToList();
                            if (check2.Count > 0)
                            {
                                var toMail2 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 3).Select(m => m.SecondApproverToList).FirstOrDefault();
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
    }
}
