using DAS.DBModels;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DAS.EntityModels.CommonEntity;

namespace DAS.DAL
{
    public class DALCommonMethod
    {
        i_facility_talContext db = new i_facility_talContext();

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALHMIWrongQty));

        public static IConfiguration configuration;

        public DALCommonMethod(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }

        #region old report updation logic

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

        //        //Taken form live chnaging to hist vignesh
        //        var SettingData = dbLoss.Tbllossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
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
        //        //Taken form live chnaging to hist vignesh
        //        var SettingData = dbLoss.Tbllossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();

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


        #region start Wo
        //Output: In Seconds

        public async Task<bool> CalWODataForYesterday(DateTime? StartDate, DateTime? EndDate,List<int> machinelist)
        {
            log.Error("CalWODataForYesterday");
            bool result = false;
            try
            {
                DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);
                //fromdate = Convert.ToDateTime(DateTime.Now.ToString("2018-05-01"));
                //todate = Convert.ToDateTime(DateTime.Now.ToString("2018-10-31"));
                if (StartDate != null && EndDate != null)
                {
                    fromdate = Convert.ToDateTime(StartDate);
                    todate = Convert.ToDateTime(EndDate);
                }

                DateTime UsedDateForExcel = Convert.ToDateTime(fromdate.ToString("yyyy-MM-dd"));
                double TotalDay = todate.Subtract(fromdate).TotalDays;

                #region
                for (int i = 0; i < TotalDay + 1; i++)
                {
                    // 2017-03-08 
                    string CorrectedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
                    //Normal WorkCenter
                    var machineData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 0 && machinelist.Contains(m.MachineId)).ToList();
                    foreach (var macrow in machineData)
                    {
                        int MachineID = macrow.MachineId;
                        //WorkOrder Data
                        try
                        {

                            ////For Testing Just Losses
                            //    int a = 0;
                            //if (a == 1)
                            //{
                            #region
                            var WODataPresent = db.Tblworeport.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).ToList();
                            if (WODataPresent.Count == 0)
                            {
                                var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
                                foreach (var hmirow in HMIData)
                                {
                                    //Constants from table

                                    int hmiid = hmirow.Hmiid;
                                    string OperatorName = hmirow.OperatorDet;
                                    string shift = hmirow.Shift;
                                    string hmiCorretedDate = hmirow.CorrectedDate;
                                    string type = hmirow.ProdFai;
                                    string program = hmirow.Project;
                                    int isHold = 0;
                                    isHold = hmirow.IsHold;
                                    DateTime StartTime = Convert.ToDateTime(hmirow.Date);
                                    DateTime EndTime = Convert.ToDateTime(hmirow.Time);
                                    //Values from Calculation
                                    double cuttingTime = 0, settingTime = 0, selfInspection = 0, idle = 0, breakdown = 0, MinorLoss = 0, SummationSCTvsPP = 0;
                                    double Blue = 0, ScrapQtyTime = 0, ReworkTime = 0;

                                    cuttingTime = await GetGreen(CorrectedDate, StartTime, EndTime, MachineID);
                                    cuttingTime = Math.Round(cuttingTime / 60, 2);
                                    settingTime = await GetSettingTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    settingTime = Math.Round(settingTime / 60, 2);
                                    selfInspection = await GetSelfInsepectionForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    selfInspection = Math.Round(selfInspection / 60, 2);
                                    double TotalLosses = await GetAllLossesTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    TotalLosses = Math.Round(TotalLosses / 60, 2);
                                    idle = TotalLosses;
                                    breakdown = await GetDownTimeBreakdownForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    breakdown = Math.Round(breakdown / 60, 2);
                                    MinorLoss = await GetMinorLossForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    MinorLoss = Math.Round(MinorLoss / 60, 2);

                                    Blue = await GetBlue(CorrectedDate, StartTime, EndTime, MachineID);
                                    Blue = Math.Round(Blue / 60, 2); bool isRework = false;
                                    isRework = hmirow.IsWorkOrder == 0 ? false : true;
                                    if (isRework)
                                    {
                                        ReworkTime = cuttingTime;
                                    }

                                    int isSingleWo = 0;
                                    isSingleWo = hmirow.IsMultiWo;

                                    if (isSingleWo == 0)
                                    {
                                        #region singleWO
                                        string SplitWO = hmirow.SplitWo;

                                        try
                                        {
                                            string PartNo = hmirow.PartNo;
                                            string WONo = hmirow.WorkOrderNo;
                                            string OpNo = hmirow.OperationNo;


                                            int targetQty = Convert.ToInt32(hmirow.TargetQty);
                                            int deliveredQty = Convert.ToInt32(hmirow.DeliveredQty);
                                            int rejectedQty = Convert.ToInt32(hmirow.RejQty);
                                            if (rejectedQty > 0)
                                            {
                                                ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
                                            }

                                            int IsPF = 0;
                                            if (hmirow.IsWorkInProgress == 1)
                                            {
                                                IsPF = 1;
                                            }

                                            //Constants From DB
                                            double stdCuttingTime = 0, stdMRWeight = 0;
                                            var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
                                            if (StdWeightTime != null)
                                            {
                                                string stdCuttingTimeString = null, stdMRWeightString = null;
                                                string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
                                                stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
                                                stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
                                                stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
                                                stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

                                                double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                                                double.TryParse(stdMRWeightString, out stdMRWeight);

                                                if (stdCuttingTimeUnitString == "Hrs")
                                                {
                                                    stdCuttingTime = stdCuttingTime * 60;
                                                }
                                                else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
                                                {
                                                    stdCuttingTime = stdCuttingTime / 60;
                                                }

                                                SummationSCTvsPP = stdCuttingTime * deliveredQty;



                                                // no need of else its already in minutes
                                            }

                                            double totalNCCuttingTime = deliveredQty * stdCuttingTime;
                                            //??
                                            string MRReason = null;

                                            double WOEfficiency = 0;
                                            if (cuttingTime != 0)
                                            {
                                                WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2) * 100;
                                                //WOEfficiency = Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100;
                                            }
                                            //Now insert into table
                                            //using (i_facility_talContext db = new i_facility_talContext())
                                            //{
                                                log.Error("WOreport inserted");
                                                Tblworeport objwo = new Tblworeport();
                                                objwo.MachineId = MachineID;
                                                objwo.Hmiid = hmiid;
                                                objwo.OperatorName = OperatorName;
                                                objwo.Shift = shift;
                                                objwo.CorrectedDate = hmiCorretedDate;
                                                objwo.PartNo = PartNo;
                                                objwo.WorkOrderNo = WONo;
                                                objwo.OpNo = OpNo;
                                                objwo.TargetQty = targetQty;
                                                objwo.DeliveredQty = deliveredQty;
                                                objwo.IsPf = IsPF;
                                                objwo.IsHold = isHold;
                                                objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
                                                objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
                                                objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
                                                objwo.Idle = (decimal)Math.Round(idle, 2);
                                                objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
                                                objwo.Type = type;
                                                objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
                                                objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
                                                objwo.Woefficiency = (decimal)WOEfficiency;
                                                objwo.RejectedQty = rejectedQty;
                                                objwo.Program = program;
                                                objwo.Mrweight = (decimal)stdMRWeight;
                                                objwo.InsertedOn = DateTime.Now;
                                                objwo.IsMultiWo = isSingleWo;
                                                objwo.MinorLoss = (decimal)Math.Round(MinorLoss, 2);
                                                objwo.SplitWo = SplitWO;
                                                objwo.Blue = (decimal)Math.Round(Blue, 2);
                                                objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
                                                objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
                                                objwo.SummationOfSctvsPp = (decimal)Math.Round(SummationSCTvsPP, 2);
                                                objwo.StartTime = StartTime;
                                                objwo.EndTime = EndTime;
                                                db.Tblworeport.Add(objwo);
                                                db.SaveChanges();
                                                log.Error("WOreport inserted success");
                                                result = true;
                                                //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
                                                //        "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,Program,MRWeight,InsertedOn,IsMultiWO,MinorLoss,SplitWO,Blue,ScrapQtyTime,ReWorkTime,SummationOfSCTvsPP,StartTime,EndTime)"
                                                //        + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "','"
                                                //        + PartNo + "','" + WONo + "','" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "','" + Math.Round(MinorLoss, 2) + "','" + SplitWO + "','" + Math.Round(Blue, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReworkTime, 2) + "','" + Math.Round(SummationSCTvsPP, 2) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");
                                           // }
                                        }
                                        catch (Exception eSingle)
                                        {
                                            result = false;
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region MultiWO
                                        var MultiWOData = db.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
                                        foreach (var multirow in MultiWOData)
                                        {
                                            string SplitWO = multirow.SplitWo;
                                            try
                                            {
                                                string PartNo = multirow.PartNo;
                                                string WONo = multirow.WorkOrder;
                                                string OpNo = multirow.OperationNo;
                                                int targetQty = Convert.ToInt32(multirow.TargetQty);
                                                int deliveredQty = Convert.ToInt32(multirow.DeliveredQty);
                                                int rejectedQty = Convert.ToInt32(multirow.ScrapQty);
                                                if (rejectedQty > 0)
                                                {
                                                    ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
                                                }

                                                int IsPF = 0;
                                                if (multirow.IsCompleted == 1)
                                                {
                                                    IsPF = 1;
                                                }
                                                //Constants From DB
                                                double stdCuttingTime = 0, stdMRWeight = 0;
                                                var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
                                                if (StdWeightTime != null)
                                                {
                                                    string stdCuttingTimeString = null, stdMRWeightString = null;
                                                    string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
                                                    stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
                                                    stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
                                                    stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
                                                    stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

                                                    double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                                                    double.TryParse(stdMRWeightString, out stdMRWeight);

                                                    if (stdCuttingTimeUnitString == "Hrs")
                                                    {
                                                        stdCuttingTime = stdCuttingTime * 60;
                                                    }
                                                    else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
                                                    {
                                                        stdCuttingTime = stdCuttingTime / 60;
                                                    }
                                                    SummationSCTvsPP = stdCuttingTime * deliveredQty;
                                                }
                                                double totalNCCuttingTime = deliveredQty * stdCuttingTime;
                                                //??
                                                string MRReason = null;

                                                double WOEfficiency = 0;
                                                if (cuttingTime != 0)
                                                {
                                                    WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2);
                                                }

                                                //Now insert into table
                                                //using (i_facility_talContext db = new i_facility_talContext())
                                                //{
                                                    try
                                                    {
                                                        log.Error("WOreport inserted in else part");
                                                        Tblworeport objwo = new Tblworeport();
                                                        objwo.MachineId = MachineID;
                                                        objwo.Hmiid = hmiid;
                                                        objwo.OperatorName = OperatorName;
                                                        objwo.Shift = shift;
                                                        objwo.CorrectedDate = hmiCorretedDate;
                                                        objwo.PartNo = PartNo;
                                                        objwo.WorkOrderNo = WONo;
                                                        objwo.OpNo = OpNo;
                                                        objwo.TargetQty = targetQty;
                                                        objwo.DeliveredQty = deliveredQty;
                                                        objwo.IsPf = IsPF;
                                                        objwo.IsHold = isHold;
                                                        objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
                                                        objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
                                                        objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
                                                        objwo.Idle = (decimal)Math.Round(idle, 2);
                                                        objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
                                                        objwo.Type = type;
                                                        objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
                                                        objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
                                                        objwo.Woefficiency = (decimal)WOEfficiency;
                                                        objwo.RejectedQty = rejectedQty;
                                                        objwo.Program = program;
                                                        objwo.Mrweight = (decimal)stdMRWeight;
                                                        objwo.InsertedOn = DateTime.Now;
                                                        objwo.IsMultiWo = isSingleWo;
                                                        objwo.MinorLoss = (decimal)Math.Round(MinorLoss, 2);
                                                        objwo.SplitWo = SplitWO;
                                                        objwo.Blue = (decimal)Math.Round(Blue, 2);
                                                        objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
                                                        objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
                                                        objwo.SummationOfSctvsPp = (decimal)Math.Round(SummationSCTvsPP, 2);
                                                        objwo.StartTime = StartTime;
                                                        objwo.EndTime = EndTime;
                                                        db.Tblworeport.Add(objwo);
                                                        db.SaveChanges();
                                                        log.Error("WOreport inserted success in else part");
                                                        result = true;
                                                        //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
                                                        //        "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,RejectedReason,Program,MRWeight,InsertedOn,IsMultiWO,MinorLoss,SplitWO,Blue,ScrapQtyTime,ReWorkTime,SummationOfSCTvsPP,StartTime,EndTime)"
                                                        //        + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "','"
                                                        //        + PartNo + "','" + WONo + "','" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + MRReason + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "','" + Math.Round(MinorLoss, 2) + "','" + SplitWO + "','" + Math.Round(Blue) + "','" + Math.Round(ScrapQtyTime) + "','" + Math.Round(ReworkTime) + "','" + Math.Round(SummationSCTvsPP) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");
                                                    }
                                                    catch (Exception eMulti)
                                                    {
                                                        result = false;
                                                    }

                                                //}
                                            }
                                            catch (Exception ex)
                                            {
                                                result = false;
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                //result = true;
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            result = false;
                        }
                        //LossesData for each WorkOrder
                        try
                        {
                            #region
                            ////Testing 
                            //MachineID = 1;
                            //CorrectedDate = "2017-03-22";

                            //var HMIData = db.tblhmiscreens.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
                            var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
                            foreach (var hmirow in HMIData)
                            {
                                int hmiid = hmirow.Hmiid;
                                var WODataPresent = db.Tblwolossess.Where(m => m.Hmiid == hmiid).ToList();
                                if (WODataPresent.Count == 0)
                                {
                                    DateTime StartTime = Convert.ToDateTime(hmirow.Date);
                                    DateTime EndTime = Convert.ToDateTime(hmirow.Time);

                                    var LossesIDs = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeId).Distinct().ToList();
                                    foreach (var loss in LossesIDs)
                                    {
                                        double duration = 0;
                                        int lossID = loss;
                                        //using (i_facility_talContext db = new i_facility_talContext())
                                        //{
                                            var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.MessageCodeId == lossID && m.DoneWithRow == 1 && (m.StartDateTime <= StartTime && ((m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime))) || (m.StartDateTime > StartTime && (m.StartDateTime < EndTime)))).ToList();

                                            foreach (var row in query2)
                                            {
                                                if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
                                                {
                                                    DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
                                                    DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
                                                    double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                                                    //Get Duration Based on start & end Time.

                                                    if (LStartDate < StartTime)
                                                    {
                                                        double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                                                        IndividualDur -= StartDurationExtra;
                                                    }
                                                    if (LEndDate > EndTime)
                                                    {
                                                        double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                                                        IndividualDur -= EndDurationExtra;
                                                    }
                                                    duration += IndividualDur;
                                                }
                                            }
                                       // }
                                        if (duration > 0)
                                        {
                                            duration = Math.Round(duration / 60, 2);
                                            //durationList.Add(new KeyValuePair<int, double>(lossID, duration));

                                            //Get Loss level, and hierarchical details
                                            int losslevel = 0, level1ID = 0, level2ID = 0;
                                            string LossName, Level1Name, Level2Name;
                                            var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == lossID).FirstOrDefault();
                                            int level = lossdata.LossCodesLevel;
                                            string losscodeName = null;

                                            #region To Get LossCode Hierarchy and Push into table
                                            if (level == 3)
                                            {
                                                int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
                                                int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2Id);
                                                var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
                                                var lossdata2 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel2ID).FirstOrDefault();
                                                losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                                Level1Name = lossdata1.LossCode;
                                                Level2Name = lossdata2.LossCode;
                                                LossName = lossdata.LossCode;

                                                //Now insert into table
                                                //using (i_facility_talContext db = new i_facility_talContext())
                                                //{
                                                    try
                                                    {
                                                        log.Error("WOlosses inserted");
                                                        Tblwolossess objwo = new Tblwolossess();
                                                        objwo.Hmiid = hmiid;
                                                        objwo.LossId = lossID;
                                                        objwo.LossName = LossName;
                                                        objwo.LossDuration = (decimal)duration;
                                                        objwo.Level = level;
                                                        objwo.LossCodeLevel1Id = lossLevel1ID;
                                                        objwo.LossCodeLevel1Name = Level1Name;
                                                        objwo.LossCodeLevel2Id = lossLevel2ID;
                                                        objwo.LossCodeLevel2Name = Level2Name;
                                                        objwo.InsertedOn = DateTime.Now;
                                                        objwo.IsDeleted = 0;
                                                        db.Tblwolossess.Add(objwo);
                                                        db.SaveChanges();
                                                        log.Error("woloss inserted success");
                                                        result = true;
                                                        //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                        //        + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,LossCodeLevel2ID,LossCodeLevel2Name,InsertedOn,IsDeleted) "
                                                        //        + " VALUES('" + hmiid + "','" + lossID + "','" + LossName + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
                                                        //        + Level1Name + "','" + lossLevel2ID + "','" + Level2Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        result = false;
                                                    }
                                               // }


                                            }
                                            else if (level == 2)
                                            {
                                                int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
                                                var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
                                                losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                                Level1Name = lossdata1.LossCode;

                                                //Now insert into table
                                                //using (i_facility_talContext db = new i_facility_talContext())
                                                //{
                                                    try
                                                    {
                                                        log.Error("wolosses inserted in else part");
                                                        Tblwolossess objwo = new Tblwolossess();
                                                        objwo.Hmiid = hmiid;
                                                        objwo.LossId = lossID;
                                                        objwo.LossName = lossdata.LossCode;
                                                        objwo.LossDuration = (decimal)duration;
                                                        objwo.Level = level;
                                                        objwo.LossCodeLevel1Id = lossLevel1ID;
                                                        objwo.LossCodeLevel1Name = Level1Name;
                                                        objwo.InsertedOn = DateTime.Now;
                                                        objwo.IsDeleted = 0;
                                                        db.Tblwolossess.Add(objwo);
                                                        db.SaveChanges();
                                                        log.Error("wolosses inserted success in else part");
                                                        result = true;
                                                        //                        SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                        //+ "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,InsertedOn,IsDeleted) "
                                                        //+ " VALUES('" + hmiid + "','" + lossID + "','" + lossdata.LossCode + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
                                                        //+ Level1Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        result = false;
                                                    }
                                               // }

                                            }
                                            else if (level == 1)
                                            {
                                                if (lossID == 999)
                                                {
                                                    losscodeName = "NoCode Entered";
                                                }
                                                else
                                                {
                                                    losscodeName = lossdata.LossCode;
                                                }
                                                //Now insert into table
                                                //using (i_facility_talContext db = new i_facility_talContext())
                                                //{
                                                    try
                                                    {
                                                        log.Error("woloss inserted for nocode");
                                                        Tblwolossess objwo = new Tblwolossess();
                                                        objwo.Hmiid = hmiid;
                                                        objwo.LossId = lossID;
                                                        objwo.LossName = lossdata.LossCode;
                                                        objwo.LossDuration = (decimal)duration;
                                                        objwo.Level = level;
                                                        objwo.InsertedOn = DateTime.Now;
                                                        objwo.IsDeleted = 0;
                                                        db.Tblwolossess.Add(objwo);
                                                        db.SaveChanges();
                                                        log.Error("woloss inserted success for nocode");
                                                        result = true;
                                                        //                        SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                        //+ "(HMIID,LossID,LossName,LossDuration,Level,InsertedOn,IsDeleted) "
                                                        //+ " VALUES('" + hmiid + "','" + lossID + "','" + losscodeName + "','" + duration + "','" + level + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0);");
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        result = false;
                                                    }
                                               // }
                                            }
                                            #endregion

                                        }
                                    }
                                }

                            }
                            //result = true;
                            #endregion
                        }
                        catch (Exception e)
                        {

                        }
                    }

                    //For Manual WorkCenters.
                    var MWCData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 1 && m.ManualWcid.HasValue).ToList();
                    foreach (var macrow in MWCData)
                    {
                        int MachineID = macrow.MachineId;
                        try
                        {
                            #region
                            var WODataPresent = db.Tblworeport.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).ToList();
                            if (WODataPresent.Count == 0)
                            {
                                var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
                                foreach (var hmirow in HMIData)
                                {
                                    //Constants from table

                                    int hmiid = hmirow.Hmiid;
                                    string OperatorName = hmirow.OperatorDet;
                                    string shift = hmirow.Shift;
                                    string hmiCorretedDate = hmirow.CorrectedDate;
                                    string type = hmirow.ProdFai;
                                    string program = hmirow.Project;
                                    int isHold = 0;
                                    isHold = hmirow.IsHold;
                                    string SplitWO = hmirow.SplitWo;
                                    int HoldID = 0; string HoldReasonID = null;
                                    try
                                    {
                                        HoldID = Convert.ToInt32(db.Tblmanuallossofentry.Where(m => m.Hmiid == hmiid).Select(m => m.MessageCodeId).FirstOrDefault());
                                    }
                                    catch (Exception e)
                                    {

                                    }
                                    if (HoldID != 0)
                                    {
                                        HoldReasonID = HoldID.ToString();
                                    }

                                    DateTime StartTime = Convert.ToDateTime(hmirow.Date);
                                    DateTime EndTime = Convert.ToDateTime(hmirow.Time);
                                    //Values from Calculation
                                    double cuttingTime = 0, settingTime = 0, selfInspection = 0, idle = 0, breakdown = 0;
                                    double Blue = 0, ScrapQtyTime = 0, ReworkTime = 0;

                                    settingTime = await GetSettingTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    settingTime = Math.Round(settingTime / 60, 2);
                                    selfInspection = await GetSelfInsepectionForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    selfInspection = Math.Round(selfInspection / 60, 2);
                                    double TotalLosses = await GetAllLossesTimeForWO(CorrectedDate, MachineID, StartTime, EndTime);
                                    TotalLosses = Math.Round(TotalLosses / 60, 2);
                                    idle = TotalLosses;
                                    breakdown = 0;

                                    var HMIIDData = db.Tblhmiscreen.Where(m => m.Hmiid == hmiid).FirstOrDefault();
                                    DateTime WOStartDateTime = Convert.ToDateTime(HMIIDData.Date);
                                    DateTime WOEndDateTime = Convert.ToDateTime(HMIIDData.Time);
                                    double TotalWODurationIsSec = WOEndDateTime.Subtract(WOStartDateTime).TotalMinutes;

                                    cuttingTime = TotalWODurationIsSec - settingTime - selfInspection;

                                    int isSingleWo = 0;
                                    isSingleWo = hmirow.IsMultiWo;

                                    try
                                    {
                                        string PartNo = hmirow.PartNo;
                                        string WONo = hmirow.WorkOrderNo;
                                        string OpNo = hmirow.OperationNo;
                                        int targetQty = Convert.ToInt32(hmirow.TargetQty);
                                        int deliveredQty = Convert.ToInt32(hmirow.DeliveredQty);
                                        int rejectedQty = Convert.ToInt32(hmirow.RejQty);
                                        int IsPF = 0;
                                        if (hmirow.IsWorkInProgress == 1)
                                        {
                                            IsPF = 1;
                                        }

                                        if (rejectedQty > 0)
                                        {
                                            ScrapQtyTime = (cuttingTime / (rejectedQty + deliveredQty)) * rejectedQty;
                                        }

                                        bool isRework = false;
                                        isRework = hmirow.IsWorkOrder == 1 ? true : false;
                                        if (isRework)
                                        {
                                            ReworkTime = cuttingTime;
                                        }

                                        //Constants From DB
                                        double stdCuttingTime = 0, stdMRWeight = 0;
                                        var StdWeightTime = db.TblmasterpartsStSw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo && m.IsDeleted == 0).FirstOrDefault();
                                        if (StdWeightTime != null)
                                        {
                                            string stdCuttingTimeString = null, stdMRWeightString = null;
                                            string stdCuttingTimeUnitString = null, stdMRWeightUnitString = null;
                                            stdCuttingTimeString = Convert.ToString(StdWeightTime.StdCuttingTime);
                                            stdMRWeightString = Convert.ToString(StdWeightTime.MaterialRemovedQty);
                                            stdCuttingTimeUnitString = Convert.ToString(StdWeightTime.StdCuttingTimeUnit);
                                            stdMRWeightUnitString = Convert.ToString(StdWeightTime.MaterialRemovedQtyUnit);

                                            double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                                            double.TryParse(stdMRWeightString, out stdMRWeight);

                                            stdCuttingTimeUnitString = StdWeightTime.StdCuttingTimeUnit;
                                            stdCuttingTimeUnitString = StdWeightTime.StdCuttingTimeUnit;

                                            if (stdCuttingTimeUnitString == "Hrs")
                                            {
                                                stdCuttingTime = stdCuttingTime * 60;
                                            }
                                            else if (stdCuttingTimeUnitString == "Sec") //Unit is Minutes
                                            {
                                                stdCuttingTime = stdCuttingTime / 60;
                                            }
                                        }
                                        double totalNCCuttingTime = deliveredQty * stdCuttingTime;
                                        //??
                                        string MRReason = null;

                                        double WOEfficiency = 0;
                                        if (cuttingTime != 0)
                                        {
                                            WOEfficiency = Math.Round((totalNCCuttingTime / cuttingTime), 2) * 100;
                                            //WOEfficiency = Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100;
                                        }
                                        //Now insert into table

                                        //using (i_facility_talContext db = new i_facility_talContext())
                                        //{
                                            try
                                            {
                                                Tblworeport objwo = new Tblworeport();
                                                objwo.MachineId = MachineID;
                                                objwo.Hmiid = hmiid;
                                                objwo.OperatorName = OperatorName;
                                                objwo.Shift = shift;
                                                objwo.CorrectedDate = hmiCorretedDate;
                                                objwo.PartNo = PartNo;
                                                objwo.WorkOrderNo = WONo;
                                                objwo.OpNo = OpNo;
                                                objwo.TargetQty = targetQty;
                                                objwo.DeliveredQty = deliveredQty;
                                                objwo.IsPf = IsPF;
                                                objwo.IsHold = isHold;
                                                objwo.CuttingTime = (decimal)Math.Round(cuttingTime, 2);
                                                objwo.SettingTime = (decimal)Math.Round(settingTime, 2);
                                                objwo.SelfInspection = (decimal)Math.Round(selfInspection, 2);
                                                objwo.Idle = (decimal)Math.Round(idle, 2);
                                                objwo.Breakdown = (decimal)Math.Round(breakdown, 2);
                                                objwo.Type = type;
                                                objwo.NccuttingTimePerPart = (decimal)stdCuttingTime;
                                                objwo.TotalNccuttingTime = (decimal)totalNCCuttingTime;
                                                objwo.Woefficiency = (decimal)WOEfficiency;
                                                objwo.RejectedQty = rejectedQty;
                                                objwo.Program = program;
                                                objwo.Mrweight = (decimal)stdMRWeight;
                                                objwo.InsertedOn = DateTime.Now;
                                                objwo.IsMultiWo = isSingleWo;
                                                objwo.IsNormalWc = 1;
                                                objwo.HoldReason = HoldReasonID;
                                                objwo.SplitWo = SplitWO;
                                                objwo.Blue = (decimal)Math.Round(Blue, 2);
                                                objwo.ScrapQtyTime = (decimal)Math.Round(ScrapQtyTime, 2);
                                                objwo.ReWorkTime = (decimal)Math.Round(ReworkTime, 2);
                                                objwo.StartTime = StartTime;
                                                objwo.EndTime = EndTime;
                                                db.Tblworeport.Add(objwo);
                                                db.SaveChanges();
                                                result = true;
                                                //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblworeport " +
                                                //    "(MachineID,HMIID,OperatorName,Shift,CorrectedDate,PartNo,WorkOrderNo,OpNo,TargetQty,DeliveredQty,IsPF,IsHold,CuttingTime,SettingTime,SelfInspection,Idle,Breakdown,Type,NCCuttingTimePerPart,TotalNCCuttingTime,WOEfficiency,RejectedQty,Program,MRWeight,InsertedOn,IsMultiWO,IsNormalWC,HoldReason,SplitWO,Blue,ScrapQtyTime,ReWorkTime,StartTime, EndTime)"
                                                //    + " VALUES('" + MachineID + "','" + hmiid + "','" + OperatorName + "','" + shift + "','" + hmiCorretedDate + "',\""
                                                //    + PartNo + "\",\"" + WONo + "\",'" + OpNo + "','" + targetQty + "','" + deliveredQty + "','" + IsPF + "','" + isHold + "','" + Math.Round(cuttingTime, 2) + "','" + Math.Round(settingTime, 2) + "','" + Math.Round(selfInspection, 2) + "','" + Math.Round(idle, 2) + "','" + Math.Round(breakdown, 2) + "','" + type + "','" + stdCuttingTime + "','" + totalNCCuttingTime + "','" + WOEfficiency + "','" + rejectedQty + "','" + program + "','" + stdMRWeight + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + isSingleWo + "',1,'" + HoldReasonID + "','" + SplitWO + "','" + Math.Round(Blue, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReworkTime, 2) + "','" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "','" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "');");

                                            }
                                            catch (Exception e)
                                            {
                                                result = false;
                                            }
                                       // }
                                    }
                                    catch (Exception eSingle)
                                    {
                                        result = false;
                                    }

                                }
                            }
                            #endregion
                        }
                        catch (Exception e)
                        {
                            result = false;
                        }

                        //LossesData for each WorkOrder
                        try
                        {
                            #region

                            var HMIData = db.Tblhmiscreen.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && (m.IsWorkInProgress == 0 || m.IsWorkInProgress == 1)).ToList();
                            foreach (var hmirow in HMIData)
                            {
                                int hmiid = hmirow.Hmiid;
                                var WODataPresent = db.Tblwolossess.Where(m => m.Hmiid == hmiid).ToList();
                                if (WODataPresent.Count == 0)
                                {
                                    DateTime StartTime = Convert.ToDateTime(hmirow.Date);
                                    DateTime EndTime = Convert.ToDateTime(hmirow.Time);

                                    var LossesIDs = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeId).Distinct().ToList();
                                    foreach (var loss in LossesIDs)
                                    {

                                        double duration = 0;
                                        int lossID = loss;
                                        //using (i_facility_talContext db = new i_facility_talContext())
                                        //{
                                            var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.MessageCodeId == lossID && m.DoneWithRow == 1 && m.StartDateTime <= StartTime && m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime) || (m.StartDateTime > StartTime && m.StartDateTime < EndTime));


                                            foreach (var row in query2)
                                            {
                                                if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
                                                {
                                                    DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
                                                    DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
                                                    double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                                                    //Get Duration Based on start & end Time.

                                                    if (LStartDate < StartTime)
                                                    {
                                                        double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                                                        IndividualDur -= StartDurationExtra;
                                                    }
                                                    if (LEndDate > EndTime)
                                                    {
                                                        double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                                                        IndividualDur -= EndDurationExtra;
                                                    }
                                                    duration += IndividualDur;
                                                }
                                            }



                                            if (duration > 0)
                                            {
                                                duration = Math.Round(duration / 60, 2);
                                                //durationList.Add(new KeyValuePair<int, double>(lossID, duration));

                                                //Get Loss level, and hierarchical details
                                                string LossName, Level1Name, Level2Name;
                                                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == lossID).FirstOrDefault();
                                                int level = lossdata.LossCodesLevel;
                                                string losscodeName = null;

                                                #region To Get LossCode Hierarchy and Push into table
                                                if (level == 3)
                                                {
                                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
                                                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2Id);
                                                    var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
                                                    var lossdata2 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel2ID).FirstOrDefault();
                                                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                                    Level1Name = lossdata1.LossCode;
                                                    Level2Name = lossdata2.LossCode;
                                                    LossName = lossdata.LossCode;

                                                    //Now insert into table
                                                    //using (i_facility_talContext db1 = new i_facility_talContext())
                                                    //{
                                                        try
                                                        {
                                                            Tblwolossess objwo = new Tblwolossess();
                                                            objwo.Hmiid = hmiid;
                                                            objwo.LossId = lossID;
                                                            objwo.LossName = LossName;
                                                            objwo.LossDuration = (decimal)duration;
                                                            objwo.Level = level;
                                                            objwo.LossCodeLevel1Id = lossLevel1ID;
                                                            objwo.LossCodeLevel1Name = Level1Name;
                                                            objwo.LossCodeLevel2Id = lossLevel2ID;
                                                            objwo.LossCodeLevel2Name = Level2Name;
                                                            objwo.InsertedOn = DateTime.Now;
                                                            objwo.IsDeleted = 0;
                                                            db.Tblwolossess.Add(objwo);
                                                            db.SaveChanges();
                                                            result = true;
                                                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                            //    + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,LossCodeLevel2ID,LossCodeLevel2Name,InsertedOn,IsDeleted) "
                                                            //    + " VALUES('" + hmiid + "','" + lossID + "','" + LossName + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
                                                            //    + Level1Name + "','" + lossLevel2ID + "','" + Level2Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");

                                                        }
                                                        catch (Exception e)
                                                        {
                                                            result = false;
                                                        }
                                                   // }

                                                }
                                                else if (level == 2)
                                                {
                                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1Id);
                                                    var lossdata1 = db.Tbllossescodes.Where(m => m.LossCodeId == lossLevel1ID).FirstOrDefault();
                                                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                                    Level1Name = lossdata1.LossCode;

                                                    //Now insert into table
                                                    //using (i_facility_talContext db1 = new i_facility_talContext())
                                                    //{
                                                        try
                                                        {
                                                            Tblwolossess objwo = new Tblwolossess();
                                                            objwo.Hmiid = hmiid;
                                                            objwo.LossId = lossID;
                                                            objwo.LossName = lossdata.LossCode;
                                                            objwo.LossDuration = (decimal)duration;
                                                            objwo.Level = level;
                                                            objwo.LossCodeLevel1Id = lossLevel1ID;
                                                            objwo.LossCodeLevel1Name = Level1Name;
                                                            objwo.InsertedOn = DateTime.Now;
                                                            objwo.IsDeleted = 0;
                                                            db.Tblwolossess.Add(objwo);
                                                            db.SaveChanges();
                                                            result = true;
                                                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                            //    + "(HMIID,LossID,LossName,LossDuration,Level,LossCodeLevel1ID,LossCodeLevel1Name,InsertedOn,IsDeleted) "
                                                            //    + " VALUES('" + hmiid + "','" + lossID + "','" + lossdata.LossCode + "','" + duration + "','" + level + "','" + lossLevel1ID + "','"
                                                            //    + Level1Name + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0)");

                                                        }
                                                        catch (Exception e)
                                                        {
                                                            result = false;
                                                        }
                                                   // }

                                                }
                                                else if (level == 1)
                                                {
                                                    if (lossID == 999)
                                                    {
                                                        losscodeName = "NoCode Entered";
                                                    }
                                                    else
                                                    {
                                                        losscodeName = lossdata.LossCode;
                                                    }
                                                    //Now insert into table
                                                    //using (i_facility_talContext db1 = new i_facility_talContext())
                                                    //{
                                                        try
                                                        {
                                                            Tblwolossess objwo = new Tblwolossess();
                                                            objwo.Hmiid = hmiid;
                                                            objwo.LossId = lossID;
                                                            objwo.LossName = lossdata.LossCode;
                                                            objwo.LossDuration = (decimal)duration;
                                                            objwo.Level = level;
                                                            objwo.InsertedOn = DateTime.Now;
                                                            objwo.IsDeleted = 0;
                                                            db.Tblwolossess.Add(objwo);
                                                            db.SaveChanges();
                                                            result = true;
                                                            //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tblwolossess "
                                                            //    + "(HMIID,LossID,LossName,LossDuration,Level,InsertedOn,IsDeleted) "
                                                            //    + " VALUES('" + hmiid + "','" + lossID + "','" + losscodeName + "','" + duration + "','" + level + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "',0);");

                                                        }
                                                        catch (Exception e)
                                                        {
                                                            result = false;
                                                        }
                                                    //}
                                                }
                                                #endregion
                                            }
                                        }
                                   // }
                                }

                                #endregion
                            }
                        }
                        catch (Exception e)
                        {
                            result = false;
                        }
                    }
                    //result = true;
                    UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                }
                #endregion
            }
            catch (Exception ex)
            {
                log.Error("Main Exception:" + ex);
            }
            return await Task.FromResult<bool>(result);

        }

        public async Task<double> GetSettingTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double settingTime = 0;
            int setupid = 0;
            string settingString = "Setup";
            var setupiddata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.MessageType.Equals(settingString, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (setupiddata != null)
            {
                setupid = setupiddata.LossCodeId;
            }
            else
            {
                return -1;
            }

            //var s = string.Join(",", products.Where(p => p.ProductType == someType).Select(p => p.ProductId.ToString()));
            // getting all setup's sublevels ids.
            var SettingIDs = db.Tbllossescodes
                                .Where(m => m.LossCodesLevel1Id == setupid)
                                .Select(m => m.LossCodeId).ToList()
                                .Distinct();
            string SettingIDsString = null;
            int j = 0;
            foreach (var row in SettingIDs)
            {
                if (j != 0)
                {
                    SettingIDsString += "," + Convert.ToInt32(row);
                }
                else
                {
                    SettingIDsString = Convert.ToInt32(row).ToString();
                }
                j++;
            }

            //using (i_facility_talContext db = new i_facility_talContext())
            //{
                var query2 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && SettingIDs.Contains(m.MessageCodeId) && m.DoneWithRow == 1 && (m.StartDateTime <= StartTime && ((m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime))) || (m.StartDateTime > StartTime && (m.StartDateTime < EndTime)))).ToList();


                foreach (var row in query2)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
                    {
                        DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
                        DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                        //Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        settingTime += IndividualDur;
                    }
               // }
            }
            return await Task.FromResult<double>(settingTime);
        }

        public async Task<double> GetSelfInsepectionForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double SelfInspectionTime = 0;
            int selfinspectId = db.Tbllossescodes.Where(m => m.LossCode == "Self inspection").Select(m => m.LossCodeId).FirstOrDefault();
            int SelfInspectionid = selfinspectId;

            //using (i_facility_talContext db = new i_facility_talContext())
            //{
                var query3 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.MessageCodeId == SelfInspectionid && m.DoneWithRow == 1 && (m.StartDateTime <= StartTime && ((m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime))) || (m.StartDateTime > StartTime && (m.StartDateTime < EndTime)))).ToList();
                foreach (var row in query3)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
                    {
                        DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
                        DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                        //Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        SelfInspectionTime += IndividualDur;
                    }
               // }
            }

            return await Task.FromResult<double>(SelfInspectionTime);
        }

        public async Task<double> GetAllLossesTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double AllLossesTime = 0;

            //using (i_facility_talContext db = new i_facility_talContext())
            //{
                var query3 = db.Tbllossofentry.Where(m => m.MachineId == MachineID && m.DoneWithRow == 1 && (m.StartDateTime <= StartTime && ((m.EndDateTime > StartTime && (m.EndDateTime < EndTime || m.EndDateTime > EndTime))) || (m.StartDateTime > StartTime && (m.StartDateTime < EndTime)))).ToList();

                foreach (var row in query3)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartDateTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndDateTime)))
                    {
                        DateTime LStartDate = Convert.ToDateTime(row.StartDateTime);
                        DateTime LEndDate = Convert.ToDateTime(row.EndDateTime);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                        //Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        AllLossesTime += IndividualDur;
                    }
               // }
            }
            return await Task.FromResult<double>(AllLossesTime);
        }

        public async Task<double> GetDownTimeBreakdownForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double BreakdownTime = 0;

            //using (i_facility_talContext db = new i_facility_talContext())
            //{
                var query3 = db.Tblbreakdown.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1 && (m.StartTime <= StartTime && ((m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime))) || (m.StartTime > StartTime && (m.StartTime < EndTime)))).ToList();

                foreach (var row in query3)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
                    {
                        DateTime LStartDate = Convert.ToDateTime(row.StartTime);
                        DateTime LEndDate = Convert.ToDateTime(row.EndTime);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;
                        // Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        BreakdownTime += IndividualDur;
                    }

               // }
            }

            return await Task.FromResult<double>(BreakdownTime);
        }

        public async Task<double> GetMinorLossForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double MinorLoss = 0;

            //using (i_facility_talContext db = new i_facility_talContext())
            //{
                var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
                  && m.ColorCode == "Yellow" && m.DurationInSec < 120 && (m.StartTime <= StartTime && ((m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime))) || (m.StartTime > StartTime && (m.StartTime < EndTime)))).ToList();

                foreach (var row in query1)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
                    {
                        DateTime LStartDate = Convert.ToDateTime(row.StartTime);
                        DateTime LEndDate = Convert.ToDateTime(row.EndTime);
                        double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                        //Get Duration Based on start & end Time.

                        if (LStartDate < StartTime)
                        {
                            double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > EndTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        MinorLoss += IndividualDur;
                    }
                //}

            }

            return await Task.FromResult<double>(MinorLoss);
        }

        public void DeletePrvDaysDataFromLiveDPS()
        {
            try
            {
                string CorrectedDate = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd");
                using (i_facility_talContext dblivedps = new i_facility_talContext())
                {
                    var liveDPSData = dblivedps.Tbllivedailyprodstatus.Where(m => m.CorrectedDate == CorrectedDate).ToList();
                    if (liveDPSData != null)
                    {
                        dblivedps.Tbllivedailyprodstatus.RemoveRange(liveDPSData);
                        dblivedps.SaveChanges();
                    }
                }
            }
            catch (Exception e)
            {

            }
        }

        ////Output: Seconds.
        public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID, int HMIID)
        {
            double SQT = 0;
            //using (i_facility_talContext dbhmi = new i_facility_talContext())
            //{
                var PartsData = db.Tblhmiscreen.Where(m => m.Hmiid == HMIID).FirstOrDefault();
                if (PartsData != null)
                {
                    int scrapQty = Convert.ToInt32(PartsData.RejQty);
                    int DeliveredQty = Convert.ToInt32(PartsData.DeliveredQty);
                    Double WODuration = await GetGreen(UsedDateForExcel, StartTime, EndTime, MachineID);
                    if ((scrapQty + DeliveredQty) == 0)
                    {
                        SQT += 0;
                    }
                    else
                    {
                        SQT += (WODuration / (scrapQty + DeliveredQty)) * scrapQty;
                    }
               // }
            }
            return await Task.FromResult<double>(SQT);
        }

        ////Output: Seconds
        public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID, int HMIID)
        {
            double SQT = 0;
            //using (i_facility_talContext dbhmi = new i_facility_talContext())
            //{
                var PartsData = db.Tblhmiscreen.Where(m => m.Hmiid == HMIID).FirstOrDefault();
                if (PartsData != null)
                {
                    int scrapQty = Convert.ToInt32(PartsData.RejQty);
                    int DeliveredQty = Convert.ToInt32(PartsData.DeliveredQty);
                    SQT = await GetGreen(UsedDateForExcel, StartTime, EndTime, MachineID);
               // }
            }
            return await Task.FromResult<double>(SQT);
        }

        ////Output: Minutes
        public async Task<double> GetSummationOfSCTvsPPForWO(int HMIID)
        {
            double SummationofTime = 0;
            //using (i_facility_talContext dbhmi = new i_facility_talContext())
            //{
                var PartsDataAll = db.Tblhmiscreen.Where(m => m.Hmiid == HMIID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
                if (PartsDataAll.Count == 0)
                {
                    //return SummationofTime;
                }
                foreach (var row in PartsDataAll)
                {
                    if (row.IsMultiWo == 0)
                    {
                        string partNo = row.PartNo;
                        string woNo = row.WorkOrderNo;
                        string opNo = row.OperationNo;
                        int DeliveredQty = 0;
                        DeliveredQty = Convert.ToInt32(row.DeliveredQty);
                        #region InnerLogic Common for both ways(HMI or tblmultiWOselection)
                        double stdCuttingTime = 0;
                        var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                        if (stdcuttingTimeData != null)
                        {
                            double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                            if (Unit == "Hrs")
                            {
                                stdCuttingTime = stdcuttingval * 60;
                            }
                            else if (Unit == "Sec") //Unit is Minutes
                            {
                                stdCuttingTime = stdcuttingval / 60;
                            }
                            else
                            {
                                stdCuttingTime = stdcuttingval;
                            }
                        }
                        #endregion
                        SummationofTime += stdCuttingTime * DeliveredQty;
                    }
                    else
                    {
                        int hmiid = row.Hmiid;
                        var multiWOData = db.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
                        foreach (var rowMulti in multiWOData)
                        {
                            string partNo = rowMulti.PartNo;
                            string opNo = rowMulti.OperationNo;
                            int DeliveredQty = 0;
                            DeliveredQty = Convert.ToInt32(rowMulti.DeliveredQty);
                            #region
                            double stdCuttingTime = 0;
                            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                            if (stdcuttingTimeData != null)
                            {
                                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                                if (Unit == "Hrs")
                                {
                                    stdCuttingTime = stdcuttingval * 60;
                                }
                                else if (Unit == "Sec") //Unit is Minutes
                                {
                                    stdCuttingTime = stdcuttingval / 60;
                                }
                                else
                                {
                                    stdCuttingTime = stdcuttingval;
                                }

                            }
                            #endregion
                            SummationofTime += stdCuttingTime * DeliveredQty;
                        }
                    }
               // }
            }
            return await Task.FromResult<double>(SummationofTime);
        }

        #endregion WO

        #region UpdateOEETable

        public async Task<bool> CalculateOEEForYesterday(DateTime? StartDate, DateTime? EndDate, List<int> machinelist)
        {
            log.Error("CalculateOEEForYesterday");
            bool result = false;
            try
            {
                //MessageBox.Show("StartTime= " + StartDate + " EndTime= " + EndDate);

                DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);

                if (StartDate != null && EndDate != null)
                {
                    fromdate = Convert.ToDateTime(StartDate);
                    todate = Convert.ToDateTime(EndDate);
                }
                //fromdate = Convert.ToDateTime(DateTime.Now.ToString("2018-05-01"));
                //todate = Convert.ToDateTime(DateTime.Now.ToString("2018-10-31"));

                //commented by V For calculating  sent date
                //fromdate = StartDate ?? DateTime.Now.AddDays(-1);
                //todate = EndDate ?? DateTime.Now.AddDays(-1);

                //DateTime fromdate = DateTime.Now.AddDays(-1), todate = DateTime.Now.AddDays(-1);
                DateTime UsedDateForExcel = Convert.ToDateTime(fromdate.ToString("yyyy-MM-dd 00:00:00"));
                double TotalDay = todate.Subtract(fromdate).TotalDays;
                #region
                for (int i = 0; i < TotalDay + 1; i++)
                {
                    //2017 - 02 - 17
                    var machineData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWc == 0 && machinelist.Contains(m.MachineId)).ToList();
                    foreach (var macrow in machineData)
                    {
                        int MachineID = macrow.MachineId;

                        try
                        {
                            log.Error("OEECal before:");
                            var OEEDataPresent = db.Tbloeedashboardvariables.Where(m => m.Wcid == MachineID && m.StartDate == UsedDateForExcel).ToList();
                            if (OEEDataPresent.Count == 0)
                            {
                                log.Error("OEECal after:" + OEEDataPresent.Count);
                                double green, red, yellow, blue, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
                                double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
                                double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;

                                MinorLosses = await GetMinorLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
                                log.Error("OEECal MinorLosses:" + MinorLosses);
                                if (MinorLosses < 0)
                                {
                                    MinorLosses = 0;
                                }
                                blue = await GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
                                log.Error("OEECal blue:" + blue);
                                green = await GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
                                log.Error("OEECal green:" + green);
                                try
                                {
                                    //Availability
                                    SettingTime = await GetSettingTime(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                                    log.Error("OEECal SettingTime:" + SettingTime);
                                    if (SettingTime < 0)
                                    {
                                        SettingTime = 0;
                                    }
                                    ROALossess = await GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROA");
                                    log.Error("OEECal ROALossess:" + ROALossess);
                                    if (ROALossess < 0)
                                    {
                                        ROALossess = 0;
                                    }
                                    DownTimeBreakdown = await GetDownTimeBreakdown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                                    log.Error("OEECal DownTimeBreakdown:" + DownTimeBreakdown);
                                    if (DownTimeBreakdown < 0)
                                    {
                                        DownTimeBreakdown = 0;
                                    }

                                    //Performance
                                    SummationOfSCTvsPP = await GetSummationOfSCTvsPP(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                                    log.Error("OEECal SummationOfSCTvsPP:" + SummationOfSCTvsPP);
                                    if (SummationOfSCTvsPP <= 0)
                                    {
                                        SummationOfSCTvsPP = 0;
                                    }

                                    //ROPLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROP");
                                }
                                catch (Exception e)
                                {
                                    result = false;
                                    log.Error("GetSummationOfSCTvsPP");
                                }

                                //Quality
                                try
                                {
                                    ScrapQtyTime = await GetScrapQtyTimeOfWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                                    log.Error("OEECal ScrapQtyTime:" + ScrapQtyTime);
                                    if (ScrapQtyTime < 0)
                                    {
                                        ScrapQtyTime = 0;
                                    }
                                    ReWOTime = await GetScrapQtyTimeOfRWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                                    log.Error("OEECal ReWOTime:" + ReWOTime);
                                    if (ReWOTime < 0)
                                    {
                                        ReWOTime = 0;
                                    }
                                }
                                catch (Exception e)
                                {
                                    result = false;
                                    log.Error("GetScrapQtyTimeOfRWO");
                                }
                                //Take care when using Available Time in Calculation of OEE and Stuff.
                                //if (TimeType == "GodHours")
                                //{
                                //    AvailableTime = AvailableTime = 24 * 60; //24Hours to Minutes;
                                //}

                                OperatingTime = green;

                                //To get Top 5 Losses for this WC
                                string todayAsCorrectedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
                                DataTable DTLosses = new DataTable();
                                DTLosses.Columns.Add("lossCodeID", typeof(int));
                                DTLosses.Columns.Add("LossDuration", typeof(int));


                                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                                //{
                                    var lossData = db.Tbllossofentry.Where(m => m.CorrectedDate == todayAsCorrectedDate && m.MachineId == MachineID).ToList();
                                    foreach (var row in lossData)
                                    {
                                        int lossCodeID = Convert.ToInt32(row.MessageCodeId);
                                        DateTime startDate = Convert.ToDateTime(row.StartDateTime);
                                        DateTime endDate = Convert.ToDateTime(row.EndDateTime);
                                        int duration = Convert.ToInt32(endDate.Subtract(startDate).TotalMinutes);

                                        DataRow dr = DTLosses.Select("lossCodeID= '" + lossCodeID + "'").FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                                        if (dr != null)
                                        {
                                            int LossDurationPrev = Convert.ToInt32(dr["LossDuration"]); //get lossduration and update it.
                                            dr["LossDuration"] = (LossDurationPrev + duration);
                                        }
                                        //}
                                        else
                                        {
                                            DTLosses.Rows.Add(lossCodeID, duration);
                                        }
                                    }
                               // }
                                DataTable DTLossesTop5 = DTLosses.Clone();
                                //get only the rows you want
                                DataRow[] results = DTLosses.Select("", "LossDuration DESC");
                                //populate new destination table
                                if (DTLosses.Rows.Count > 0)
                                {
                                    int num = DTLosses.Rows.Count;
                                    for (var iDT = 0; iDT < num; iDT++)
                                    {
                                        if (results[iDT] != null)
                                        {
                                            DTLossesTop5.ImportRow(results[iDT]);
                                        }
                                        else
                                        {
                                            DTLossesTop5.Rows.Add(0, 0);
                                        }
                                        if (iDT == 4)
                                        {
                                            break;
                                        }
                                    }
                                    if (num < 5)
                                    {
                                        for (var iDT = num; iDT < 5; iDT++)
                                        {
                                            DTLossesTop5.Rows.Add(0, 0);
                                        }
                                    }
                                }
                                else
                                {
                                    for (var iDT = 0; iDT < 5; iDT++)
                                    {
                                        DTLossesTop5.Rows.Add(0, 0);
                                    }
                                }
                                ////Gather LossValues
                                string lossCode1, lossCode2, lossCode3, lossCode4, lossCode5 = null;
                                int lossCodeVal1, lossCodeVal2, lossCodeVal3, lossCodeVal4, lossCodeVal5 = 0;

                                lossCode1 = Convert.ToString(DTLossesTop5.Rows[0][0]);
                                lossCode2 = Convert.ToString(DTLossesTop5.Rows[1][0]);
                                lossCode3 = Convert.ToString(DTLossesTop5.Rows[2][0]);
                                lossCode4 = Convert.ToString(DTLossesTop5.Rows[3][0]);
                                lossCode5 = Convert.ToString(DTLossesTop5.Rows[4][0]);
                                lossCodeVal1 = Convert.ToInt32(DTLossesTop5.Rows[0][1]);
                                lossCodeVal2 = Convert.ToInt32(DTLossesTop5.Rows[1][1]);
                                lossCodeVal3 = Convert.ToInt32(DTLossesTop5.Rows[2][1]);
                                lossCodeVal4 = Convert.ToInt32(DTLossesTop5.Rows[3][1]);
                                lossCodeVal5 = Convert.ToInt32(DTLossesTop5.Rows[4][1]);

                                //Gather Plant, Shop, Cell for WC.

                                //int PlantID = 0, ShopID = 0, CellID = 0;
                                string PlantIDS = null, ShopIDS = null, CellIDS = null;
                                int value;
                                var WCData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == MachineID).FirstOrDefault();
                                string TempVal = WCData.PlantId.ToString();
                                if (int.TryParse(TempVal, out value))
                                {
                                    PlantIDS = value.ToString();
                                }

                                TempVal = WCData.ShopId.ToString();
                                if (int.TryParse(TempVal, out value))
                                {
                                    ShopIDS = value.ToString();
                                }

                                TempVal = WCData.CellId.ToString();
                                if (int.TryParse(TempVal, out value))
                                {
                                    CellIDS = value.ToString();
                                }

                                // Now insert into table
                                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                                //{
                                    try
                                    {
                                        log.Error("OeeDashboard insertion");
                                        Tbloeedashboardvariables objoee = new Tbloeedashboardvariables();
                                        objoee.PlantId = Convert.ToInt32(PlantIDS);
                                        objoee.ShopId = Convert.ToInt32(ShopIDS);
                                        objoee.CellId = Convert.ToInt32(CellIDS);
                                        objoee.Wcid = Convert.ToInt32(MachineID);
                                        objoee.StartDate = UsedDateForExcel;
                                        objoee.EndDate = UsedDateForExcel;
                                        objoee.MinorLosses = Math.Round(MinorLosses / 60, 2);
                                        objoee.Blue = Math.Round(blue / 60, 2);
                                        objoee.Green = Math.Round(green / 60, 2);
                                        objoee.SettingTime = Math.Round(SettingTime, 2);
                                        objoee.Roalossess = Math.Round(ROALossess / 60, 2);
                                        objoee.DownTimeBreakdown = Math.Round(DownTimeBreakdown, 2);
                                        objoee.SummationOfSctvsPp = Math.Round(SummationOfSCTvsPP, 2);
                                        objoee.ScrapQtyTime = Math.Round(ScrapQtyTime, 2);
                                        objoee.ReWotime = Math.Round(ReWOTime, 2);
                                        objoee.Loss1Name = lossCode1;
                                        objoee.Loss1Value = lossCodeVal1;
                                        objoee.Loss2Name = lossCode2;
                                        objoee.Loss2Value = lossCodeVal2;
                                        objoee.Loss3Name = lossCode3;
                                        objoee.Loss3Value = lossCodeVal3;
                                        objoee.Loss4Name = lossCode4;
                                        objoee.Loss4Value = lossCodeVal4;
                                        objoee.Loss5Name = lossCode5;
                                        objoee.Loss5Value = lossCodeVal5;
                                        objoee.CreatedOn = DateTime.Now;
                                        objoee.CreatedBy = 1;
                                        objoee.IsDeleted = 0;
                                        db.Tbloeedashboardvariables.Add(objoee);
                                        db.SaveChanges();
                                        result = true;
                                        log.Error("OeeDashboard insertion succeed");

                                        //SqlCommand cmdInsertRows = new SqlCommand("INSERT INTO [i_facility_tsal].[dbo].tbloeedashboardvariables (PlantID,ShopID,CellID,WCID,StartDate,EndDate,MinorLosses,Blue,Green,SettingTime,ROALossess,DownTimeBreakdown,SummationOfSCTvsPP,ScrapQtyTime,ReWOTime,Loss1Name,Loss1Value,Loss2Name,Loss2Value,Loss3Name,Loss3Value,Loss4Name,Loss4Value,Loss5Name,Loss5Value,CreatedOn,CreatedBy,IsDeleted)VALUES('" + PlantIDS + "','" + ShopIDS + "','" + CellIDS + "','" + MachineID + "','" + UsedDateForExcel.ToString("yyyy-MM-dd") + "','" + UsedDateForExcel.ToString("yyyy-MM-dd") + "','" + Math.Round(MinorLosses / 60, 2) + "','" + Math.Round(blue / 60, 2) + "','" + Math.Round(green / 60, 2) + "','" + Math.Round(SettingTime, 2) + "','" + Math.Round(ROALossess / 60, 2) + "','" + Math.Round(DownTimeBreakdown, 2) + "','" + Math.Round(SummationOfSCTvsPP, 2) + "','" + Math.Round(ScrapQtyTime, 2) + "','" + Math.Round(ReWOTime, 2) + "','" + lossCode1 + "','" + lossCodeVal1 + "','" + lossCode2 + "','" + lossCodeVal2 + "','" + lossCode3 + "','" + lossCodeVal3 + "','" + lossCode4 + "','" + lossCodeVal4 + "','" + lossCode5 + "','" + lossCodeVal5 + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + 1 + "','" + 0 + "');");

                                    }
                                    catch (Exception e)
                                    {
                                        result = false;
                                        log.Error("Exception" + e);
                                    }
                                    //finally
                                    //{
                                    //    mcInsertRows.close();
                                    //}
                                //}
                            }
                        }
                        catch (Exception e)
                        {
                            result = false;
                            log.Error("After Exception catch");
                            //IntoFile("MacID: " + MachineID + e.ToString());
                        }
                    }
                    //result = true;
                    UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                }
                #endregion
             
            }
            catch(Exception ex)
            {
                log.Error("Main Exception:" + ex);
            }
            return await Task.FromResult<bool>(result);

        }

        public async Task<double> GetMinorLosses(string CorrectedDate, int MachineID, string Colour)
        {
            double minorloss = 0;
            try
            {

                DateTime currentdate = Convert.ToDateTime(CorrectedDate);
                string dateString = currentdate.ToString("yyyy-MM-dd");

                
                #region commented
                //int count = 0;
                //var Data = db.tbldailyprodstatus.Where(m => m.IsDeleted == 0 && m.MachineId == MachineID && m.CorrectedDate == CorrectedDate).OrderBy(m => m.StartTime).ToList();
                //foreach (var row in Data)
                //{
                //    if (row.ColorCode == "yellow")
                //    {
                //        count++;
                //    }
                //    else
                //    {
                //        if (count > 0 && count < 2)
                //        {
                //            minorloss += count;
                //            count = 0;

                //        }
                //        count = 0;
                //    }
                //}

                #endregion
                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                //{
                    var MinorLossSummation = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == dateString && m.ColorCode == Colour && m.DurationInSec < 120 && m.IsCompleted == 1).Sum(m => m.DurationInSec);
                    minorloss = await Task.FromResult<double>(Convert.ToDouble(MinorLossSummation));
                //}
            }
            catch (Exception ex)
            {
                log.Error("MinorLoss" + ex);
            }
            return minorloss;
        }
        public async Task<double> GetOPIDleBreakDown(string CorrectedDate, int MachineID, string Colour)
        {
            double count = 0;
            try
            {


                DateTime currentdate = Convert.ToDateTime(CorrectedDate);
                string datetime = currentdate.ToString("yyyy-MM-dd");


                //MsqlConnection mc = new MsqlConnection();
                //mc.open();
                ////operating
                //mc.open();
                //String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
                //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                //DataTable OP = new DataTable();
                //da1.Fill(OP);
                //mc.close();
                //if (OP.Rows.Count != 0)
                //{
                //    count[0] = Convert.ToInt32(OP.Rows[0][0]);
                //}

                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                //{
                    var blah = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == CorrectedDate && m.ColorCode == Colour).Sum(m => m.DurationInSec);
                    count = await Task.FromResult<double>(Convert.ToDouble(blah));
                //}
            }
            catch(Exception ex)
            {
                log.Error("IdleBreakdown" + ex);
            }
            return count;
        }

        public async Task<double> GetSettingTime(string UsedDateForExcel, int MachineID)
        {
            double settingTime = 0;
            try
            {
                int setupid = 0;
                string settingString = "Setup";
                var setupiddata = db.Tbllossescodes.Where(m => m.MessageType.Contains(settingString)).FirstOrDefault();
                if (setupiddata != null)
                {
                    setupid = setupiddata.LossCodeId;
                }
                else
                {
                    //Session["Error"] = "Unable to get Setup's ID";
                    return -1;
                }
                // getting all setup's sublevels ids.
                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                //{
                    var SettingIDs = db.Tbllossescodes.Where(m => m.LossCodesLevel1Id == setupid || m.LossCodesLevel2Id == setupid).Select(m => m.LossCodeId).ToList();


                    //settingTime = (from row in db.tbllivelossofenties
                    //where row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );

                    //Taken form live chnaging to hist vignesh
                    var SettingData = db.Tbllossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
                    foreach (var row in SettingData)
                    {
                        DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                        DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                        settingTime += endTime.Subtract(startTime).TotalMinutes;
                    }
                //}
            }
            catch(Exception ex)
            {
                log.Error("Setting" + ex);
            }
            return await Task.FromResult<double>(settingTime);
        }
        public async Task<double> GetDownTimeLosses(string UsedDateForExcel, int MachineID, string contribute)
        {
            double LossTime = 0;
            try
            {
                //string contribute = "ROA";
                // getting all ROA sublevels ids.Only those of IDLE.

                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                //{
                    var SettingIDs = db.Tbllossescodes.Where(m => m.ContributeTo == contribute && (m.MessageType != "PM" || m.MessageType != "BREAKDOWN")).Select(m => m.LossCodeId).ToList();
                    //Taken form live chnaging to hist vignesh
                    var SettingData = db.Tbllossofentry.Where(m => SettingIDs.Contains(m.MessageCodeId) && m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();

                    var LossDuration = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.IsCompleted == 1 && m.DurationInSec > 120 && m.ColorCode == "YELLOW").Sum(m => m.DurationInSec);

                    foreach (var row in SettingData)
                    {
                        DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                        DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                        LossTime += endTime.Subtract(startTime).TotalMinutes;
                    }
                    try
                    {
                        LossTime = (int)LossDuration;
                    }
                    catch { }
               // }
            }
            catch(Exception ex)
            {
                log.Error("DownTimeLoss" + ex);
            }

            return await Task.FromResult<double>(LossTime);
        }
        public async Task<double> GetDownTimeBreakdown(string UsedDateForExcel, int MachineID)
        {
            double LossTime = 0;
            try
            {
                if (MachineID == 18)
                {
                }

                //using (i_facility_talContext dbLoss = new i_facility_talContext())
                //{
                    var BreakdownData = db.Tblbreakdown.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
                    foreach (var row in BreakdownData)
                    {
                        if ((Convert.ToString(row.EndTime) == null) || row.EndTime == null)
                        {
                            //do nothing
                        }
                        else
                        {
                            DateTime startTime = Convert.ToDateTime(row.StartTime);
                            DateTime endTime = Convert.ToDateTime(row.EndTime);
                            LossTime += endTime.Subtract(startTime).TotalMinutes;
                        }
                    }
                //}
            }
            catch(Exception ex)
            {
                log.Error("DownTimeBreakdown" + ex);
            }
            return await Task.FromResult<double>(LossTime);
        }

        public async Task<double> GetSummationOfSCTvsPP(string UsedDateForExcel, int MachineID)
        {
            double SummationofTime = 0;
            //UsedDateForExcel = "2018-12-01";

            #region OLD 2017-02-10
            //var PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
            //if (PartsData.Count == 0)
            //{
            //    //return -1;
            //}
            //foreach (var row in PartsData)
            //{
            //    string partno = row.PartNo;
            //    string operationno = row.OperationNo;
            //    int totalpartproduced = Convert.ToInt32(row.DeliveredQty) + Convert.ToInt32(row.RejQty);
            //    Double stdCuttingTime = 0;
            //    var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
            //    if (stdcuttingTimeData != null)
            //    {
            //        string stdcuttingvalString = Convert.ToString(stdcuttingTimeData.StdCuttingTime);
            //        Double stdcuttingval = 0;
            //        if (double.TryParse(stdcuttingvalString, out stdcuttingval))
            //        {
            //            stdcuttingval = stdcuttingval;
            //        }

            //        string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //        if (Unit == "Hrs")
            //        {
            //            stdCuttingTime = stdcuttingval * 60;
            //        }
            //        else //Unit is Minutes
            //        {
            //            stdCuttingTime = stdcuttingval;
            //        }
            //    }
            //    SummationofTime += stdCuttingTime * totalpartproduced;
            //}
            ////To Extract MultiWorkOrder Cutting Time
            //PartsData = db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWO == 1 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
            //if (PartsData.Count == 0)
            //{
            //    return SummationofTime;
            //}
            //foreach (var row in PartsData)
            //{
            //    int HMIID = row.HMIID;

            //    var DataInMultiwoSelection = db.tbl_multiwoselection.Where(m => m.HMIID == HMIID).ToList();
            //    foreach (var rowData in DataInMultiwoSelection)
            //    {
            //        string partno = rowData.PartNo;
            //        string operationno = rowData.OperationNo;
            //        int totalpartproduced = Convert.ToInt32(rowData.DeliveredQty) + Convert.ToInt32(rowData.ScrapQty);
            //        int stdCuttingTime = 0;
            //        var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == operationno && m.PartNo == partno).FirstOrDefault();
            //        if (stdcuttingTimeData != null)
            //        {
            //            int stdcuttingval = Convert.ToInt32(stdcuttingTimeData.StdCuttingTime);
            //            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //            if (Unit == "Hrs")
            //            {
            //                stdCuttingTime = stdcuttingval * 60;
            //            }
            //            else //Unit is Minutes
            //            {
            //                stdCuttingTime = stdcuttingval;
            //            }
            //        }
            //        SummationofTime += stdCuttingTime * totalpartproduced;
            //    }
            //}

            #endregion

            #region OLD 2017-02-10
            //List<string> OccuredWOs = new List<string>();
            ////To Extract Single WorkOrder Cutting Time
            //using (i_facility_talContext dbhmi = new i_facility_talContext())
            //{
            //    var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.Hmiid).ToList();
            //    if (PartsDataAll.Count == 0)
            //    {
            //        //return SummationofTime;
            //    }
            //    foreach (var row in PartsDataAll)
            //    {
            //        string partNo = row.PartNo;
            //        string woNo = row.Work_Order_No;
            //        string opNo = row.OperationNo;

            //        string occuredwo = partNo + "," + woNo + "," + opNo;
            //        if (!OccuredWOs.Contains(occuredwo))
            //        {
            //            OccuredWOs.Add(occuredwo);
            //            var PartsData = dbhmi.Tblhmiscreen.
            //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0
            //                    && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)
            //                    && m.WorkOrderNo == woNo && m.PartNo == partNo && m.OperationNo == opNo).
            //                    OrderByDescending(m => m.Hmiid).ToList();

            //            int totalpartproduced = 0;
            //            int ProcessQty = 0, DeliveredQty = 0;
            //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

            //            #region new code

            //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
            //            int isHMIFirst = 2; //default NO History for that wo,pn,on

            //            var mulitwoData = dbhmi.TblMultiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWoid).Take(1).ToList();
            //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.IsWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

            //            //Note: we are in this loop => hmiscreen table data is Available

            //            if (mulitwoData.Count > 0)
            //            {
            //                isHMIFirst = 1;
            //            }
            //            else if (PartsData.Count > 0)
            //            {
            //                isHMIFirst = 0;
            //            }
            //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we both Dates now check for greatest amongst
            //            {
            //                int hmiIDFromMulitWO = row.HMIID;
            //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
            //                                                              where r.HMIID == hmiIDFromMulitWO
            //                                                              select r.Time
            //                                                              );
            //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

            //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
            //                {
            //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
            //                }
            //                else
            //                {
            //                    isHMIFirst = 0;
            //                }
            //            }
            //            if (isHMIFirst == 1)
            //            {
            //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);

            //            }
            //            else if (isHMIFirst == 0)//Take Data from HMI
            //            {
            //                string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(PartsData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }

            //            #endregion

            //            //totalpartproduced = DeliveredQty + ProcessQty;
            //            totalpartproduced = DeliveredQty;

            //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

            //            double stdCuttingTime = 0;
            //            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
            //            if (stdcuttingTimeData != null)
            //            {
            //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
            //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //                if (Unit == "Hrs")
            //                {
            //                    stdCuttingTime = stdcuttingval * 60;
            //                }
            //                else //Unit is Minutes
            //                {
            //                    stdCuttingTime = stdcuttingval;
            //                }
            //            }
            //            #endregion

            //            SummationofTime += stdCuttingTime * totalpartproduced;
            //        }
            //    }
            //}
            ////To Extract Multi WorkOrder Cutting Time
            //using (i_facility_talContext dbhmi = new i_facility_talContext())
            //{
            //    var PartsDataAll = dbhmi.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 1 && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).ToList();
            //    if (PartsDataAll.Count == 0)
            //    {
            //        //return SummationofTime;
            //    }
            //    foreach (var row in PartsDataAll)
            //    {
            //        string partNo = row.PartNo;
            //        string woNo = row.WorkOrderNo;
            //        string opNo = row.OperationNo;

            //        string occuredwo = partNo + "," + woNo + "," + opNo;
            //        if (!OccuredWOs.Contains(occuredwo))
            //        {
            //            OccuredWOs.Add(occuredwo);
            //            var PartsData = dbhmi.Tblhmiscreen.
            //                Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsMultiWo == 0
            //                    && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)
            //                    && m.WorkOrderNo == woNo && m.PartNo == partNo && m.OperationNo == opNo).
            //                    OrderByDescending(m => m.Hmiid).ToList();

            //            int totalpartproduced = 0;
            //            int ProcessQty = 0, DeliveredQty = 0;
            //            //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

            //            #region new code

            //            //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
            //            int isHMIFirst = 2; //default NO History for that wo,pn,on

            //            var mulitwoData = dbhmi.TblMultiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWoid).Take(1).ToList();
            //            //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.IsWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

            //            //Note: we are in this loop => hmiscreen table data is Available

            //            if (mulitwoData.Count > 0)
            //            {
            //                isHMIFirst = 1;
            //            }
            //            else if (PartsData.Count > 0)
            //            {
            //                isHMIFirst = 0;
            //            }
            //            else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we have both Dates now check for greatest amongst
            //            {
            //                int hmiIDFromMulitWO = row.Hmiid;
            //                DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
            //                                                              where r.HMIID == hmiIDFromMulitWO
            //                                                              select r.Time
            //                                                              );
            //                DateTime hmiDateTime = Convert.ToDateTime(row.Time);

            //                if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
            //                {
            //                    isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
            //                }
            //                else
            //                {
            //                    isHMIFirst = 0;
            //                }
            //            }

            //            if (isHMIFirst == 1)
            //            {
            //                string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(mulitwoData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }
            //            else if (isHMIFirst == 0) //Take Data from HMI
            //            {
            //                string delivString = Convert.ToString(PartsData[0].DeliveredQty);
            //                int.TryParse(delivString, out DeliveredQty);
            //                string processString = Convert.ToString(PartsData[0].ProcessQty);
            //                int.TryParse(processString, out ProcessQty);
            //            }

            //            #endregion

            //            //totalpartproduced = DeliveredQty + ProcessQty;
            //            totalpartproduced = DeliveredQty;
            //            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

            //            double stdCuttingTime = 0;
            //            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
            //            if (stdcuttingTimeData != null)
            //            {
            //                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
            //                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
            //                if (Unit == "Hrs")
            //                {
            //                    stdCuttingTime = stdcuttingval * 60;
            //                }
            //                else //Unit is Minutes
            //                {
            //                    stdCuttingTime = stdcuttingval;
            //                }
            //            }
            //            #endregion

            //            SummationofTime += stdCuttingTime * totalpartproduced;
            //        }
            //    }
            //}
            #endregion

            try
            {
                //new Code 2017-03-08
                //using (i_facility_talContext dbhmi = new i_facility_talContext())
                //{
                    var PartsDataAll = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && m.IsWorkOrder == 0 && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0)).OrderByDescending(m => m.PartNo).ThenByDescending(m => m.OperationNo).ToList();
                    if (PartsDataAll.Count == 0)
                    {
                        return SummationofTime;
                    }
                    foreach (var row in PartsDataAll)
                    {
                        if (row.IsMultiWo == 0)
                        {
                            string partNo = row.PartNo;
                            string woNo = row.WorkOrderNo;
                            string opNo = row.OperationNo;
                            int DeliveredQty = 0;
                            DeliveredQty = Convert.ToInt32(row.DeliveredQty);
                            #region InnerLogic Common for both ways(HMI or tblmultiWOselection)
                            double stdCuttingTime = 0;
                            var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                            if (stdcuttingTimeData != null)
                            {
                                double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                                string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                                if (Unit == "Hrs")
                                {
                                    stdCuttingTime = stdcuttingval * 60;
                                }
                                else if (Unit == "Sec") //Unit is Minutes
                                {
                                    stdCuttingTime = stdcuttingval / 60;
                                }
                                else
                                {
                                    stdCuttingTime = stdcuttingval;
                                }
                                //no need of else , its already in minutes
                            }
                            #endregion
                            //MessageBox.Show("CuttingTime " + stdCuttingTime + " DeliveredQty " + DeliveredQty);
                            SummationofTime += stdCuttingTime * DeliveredQty;
                            //MessageBox.Show("Single" + SummationofTime);
                        }
                        else
                        {
                            int hmiid = row.Hmiid;
                            var multiWOData = db.TblMultiwoselection.Where(m => m.Hmiid == hmiid).ToList();
                            foreach (var rowMulti in multiWOData)
                            {
                                string partNo = rowMulti.PartNo;
                                string opNo = rowMulti.OperationNo;
                                int DeliveredQty = 0;
                                DeliveredQty = Convert.ToInt32(rowMulti.DeliveredQty);
                                #region
                                double stdCuttingTime = 0;
                                var stdcuttingTimeData = db.TblmasterpartsStSw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                                if (stdcuttingTimeData != null)
                                {
                                    double stdcuttingval = Convert.ToDouble(stdcuttingTimeData.StdCuttingTime);
                                    string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                                    if (Unit == "Hrs")
                                    {
                                        stdCuttingTime = stdcuttingval * 60;
                                    }
                                    else if (Unit == "Sec") //Unit is Minutes
                                    {
                                        stdCuttingTime = stdcuttingval / 60;
                                    }
                                    else
                                    {
                                        stdCuttingTime = stdcuttingval;
                                    }

                                }
                                #endregion
                                //MessageBox.Show("CuttingTime " + stdCuttingTime + " DeliveredQty " + DeliveredQty);
                                SummationofTime += stdCuttingTime * DeliveredQty;
                                //MessageBox.Show("Multi" + SummationofTime);
                            }
                        }
                        //MessageBox.Show("" + SummationofTime);
                    }
               // }
            }
            catch(Exception ex)
            {
                log.Error("GetSummationOfSCTvsPP" + ex);
            }
            return await Task.FromResult<double>(SummationofTime);
        }

        //Output in Seconds
        public async Task<double> GetGreen(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        {
            double settingTime = 0;
            try
            {
                string stTime = StartTime.ToString("yyyy-MM-dd HH:mm:ss");
                //using (i_facility_talContext db = new i_facility_talContext())
                //{
                    var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
                      && m.ColorCode == "green" && (m.StartTime <= StartTime && ((m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime))) || (m.StartTime > StartTime && (m.StartTime < EndTime)))).ToList();
                    foreach (var row in query1)
                    {
                        if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
                        {
                            DateTime LStartDate = Convert.ToDateTime(row.StartTime);
                            DateTime LEndDate = Convert.ToDateTime(row.EndTime);
                            double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                            //Get Duration Based on start & end Time.

                            if (LStartDate < StartTime)
                            {
                                double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                                IndividualDur -= StartDurationExtra;
                            }
                            if (LEndDate > EndTime)
                            {
                                double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                                IndividualDur -= EndDurationExtra;
                            }
                            settingTime += IndividualDur;
                        }
                    }
                //}
            }
            catch(Exception ex)
            {
                log.Error("GetGreen " + ex);
            }
            return await Task.FromResult<double>(settingTime);
        }

        public async Task<double> GetBlue(string UsedDateForExcel, DateTime StartTime, DateTime EndTime, int MachineID)
        {
            double settingTime = 0;
            try
            {
                //using (i_facility_talContext db = new i_facility_talContext())
                //{
                    var query1 = db.Tblmode.Where(m => m.MachineId == MachineID && m.CorrectedDate == UsedDateForExcel
                      && m.ColorCode == "Blue" && (m.StartTime <= StartTime && ((m.EndTime > StartTime && (m.EndTime < EndTime || m.EndTime > EndTime))) || (m.StartTime > StartTime && (m.StartTime < EndTime)))).ToList();

                    foreach (var row in query1)
                    {
                        if (!string.IsNullOrEmpty(Convert.ToString(row.StartTime)) && !string.IsNullOrEmpty(Convert.ToString(row.EndTime)))
                        {
                            DateTime LStartDate = Convert.ToDateTime(row.StartTime);
                            DateTime LEndDate = Convert.ToDateTime(row.EndTime);
                            double IndividualDur = LEndDate.Subtract(LStartDate).TotalSeconds;

                            //Get Duration Based on start & end Time.

                            if (LStartDate < StartTime)
                            {
                                double StartDurationExtra = StartTime.Subtract(LStartDate).TotalSeconds;
                                IndividualDur -= StartDurationExtra;
                            }
                            if (LEndDate > EndTime)
                            {
                                double EndDurationExtra = LEndDate.Subtract(EndTime).TotalSeconds;
                                IndividualDur -= EndDurationExtra;
                            }
                            settingTime += IndividualDur;
                        }
                    }
                //}
            }
            catch(Exception ex)
            {
                log.Error("GetBlue" + ex);
            }
            return await Task.FromResult<double>(settingTime);
        }

        public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            try
            {
                //using (i_facility_talContext dbhmi = new i_facility_talContext())
                //{
                    var PartsData = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 0).ToList();
                    foreach (var row in PartsData)
                    {
                        string partno = row.PartNo;
                        string operationno = row.OperationNo;
                        int scrapQty = 0;
                        int DeliveredQty = 0;
                        string scrapQtyString = Convert.ToString(row.RejQty);
                        string DeliveredQtyString = Convert.ToString(row.DeliveredQty);
                        string x = scrapQtyString;
                        int value;
                        if (int.TryParse(x, out value))
                        {
                            scrapQty = value;
                        }
                        x = DeliveredQtyString;
                        if (int.TryParse(x, out value))
                        {
                            DeliveredQty = value;
                        }

                        DateTime startTime = Convert.ToDateTime(row.Date);
                        DateTime endTime = Convert.ToDateTime(row.Time);
                        //Double WODuration = endTimeTemp.Subtract(startTime).TotalMinutes;
                        Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                        if ((scrapQty + DeliveredQty) == 0)
                        {
                            SQT += 0;
                        }
                        else
                        {
                            SQT += ((WODuration / 60) / (scrapQty + DeliveredQty)) * scrapQty;
                        }
                    }
               // }
            }
            catch(Exception ex)
            {
                log.Error("GetScrapQtyTimeOfWO" + ex);
            }
            return await Task.FromResult<double>(SQT);
        }

        //GOD
        public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            try
            {
                //using (i_facility_talContext dbhmi = new i_facility_talContext())
                //{
                    var PartsData = db.Tblhmiscreen.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineId == MachineID && (m.IsWorkInProgress == 1 || m.IsWorkInProgress == 0) && m.IsWorkOrder == 1).ToList();
                    foreach (var row in PartsData)
                    {
                        string partno = row.PartNo;
                        string operationno = row.OperationNo;
                        int scrapQty = Convert.ToInt32(row.RejQty);
                        int DeliveredQty = Convert.ToInt32(row.DeliveredQty);
                        DateTime startTime = Convert.ToDateTime(row.Date);
                        DateTime endTime = Convert.ToDateTime(row.Time);
                        Double WODuration = await GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                        //Double WODuration = endTime.Subtract(startTime).TotalMinutes;
                        ////For Availability Loss
                        //double Settingtime = GetSetupForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                        //double green = GetOT(UsedDateForExcel, startTime, endTime, MachineID);
                        //double DownTime = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROA");
                        //double BreakdownTime = GetBreakDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                        //double AL = DownTime + BreakdownTime + Settingtime;

                        ////For Performance Loss
                        //double downtimeROP = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROP");
                        //double minorlossWO = GetMinorLossForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "yellow");
                        //double PL = downtimeROP + minorlossWO;

                        SQT += (WODuration / 60);
                    }
               // }
            }
            catch(Exception ex)
            {
                log.Error("GetScrapQtyTimeOfRWO" + ex);
            }
            return await Task.FromResult<double>(SQT);
        }

        #endregion End Oee Report


    }
}
