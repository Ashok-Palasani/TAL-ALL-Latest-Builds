using DAS.DBModels;
using System;
using System.Collections.Generic;
using System.Text;
using DAS.Interface;
using DAS.EntityModels;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using static DAS.EntityModels.BatchProcessingEntity;
using static DAS.EntityModels.CommonEntity;
using DAS.DAL.Resource;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Extensions.Configuration;
using DAS.DAL.Helpers;
using Microsoft.Extensions.Options;
using System.IO;
using static DAS.EntityModels.CommonResponseWithMachineName;

namespace DAS.DAL
{
    public class DALBatch : IBatchProcess



    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALBatch));
        public static IConfiguration configuration;
        private readonly AppSettings appSettings;

        public DALBatch(i_facility_talContext _db, IConfiguration _config, IOptions<AppSettings> _appSettings)
        {
            db = _db;
            configuration = _config;
            appSettings = _appSettings.Value;
        }

        // Getting the all the data based on iscompleted=0 from tblddl
        public DDLCommonResponseForBatch GetDDLLists(DDLListForBatch data)
        {

            DDLCommonResponseForBatch obj = new DDLCommonResponseForBatch();
            int machineId = data.machineId;
            int takeValue = data.takeValue;
            int skipeValue = data.skipeValue;
            try
            {
                var machineName = db.Tblmachinedetails.Where(m => m.MachineId == machineId).Select(m => new { m.MachineInvNo, m.MachineDispName }).FirstOrDefault();
                //getting the connection string from app string.json
                string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;

                int count = 0;
                List<EDDLList> listEDDLList = new List<EDDLList>();
                try
                {
                    DataSet ds = new DataSet();
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    SqlCommand SqlCommand = new SqlCommand("GetDDLForBatchProcess", sqlConnection);
                    SqlCommand.CommandType = CommandType.StoredProcedure;
                    SqlCommand.Parameters.AddWithValue("@takeValue", takeValue);
                    SqlCommand.Parameters.AddWithValue("@skipeValue", skipeValue);

                    string machineInvNo = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => x.MachineInvNo).FirstOrDefault();
                    count = db.Tblddl.Where(m => m.WorkCenter == machineInvNo).Count();
                    SqlCommand.Parameters.AddWithValue("@workCenter", machineInvNo);

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
                                    objEDDLList.partNo = (ds.Tables[0].Rows[i]["PartName"]).ToString();
                                    objEDDLList.project = (ds.Tables[0].Rows[i]["Project"]).ToString();
                                    objEDDLList.splitWO = (ds.Tables[0].Rows[i]["SplitWO"]).ToString();
                                    objEDDLList.targetQty = (ds.Tables[0].Rows[i]["TargetQty"]).ToString();
                                    objEDDLList.woNo = (ds.Tables[0].Rows[i]["WorkOrder"]).ToString();
                                    objEDDLList.pcpNo = (ds.Tables[0].Rows[i]["pcpNo"]).ToString();
                                    objEDDLList.slno = slno;
                                    listEDDLList.Add(objEDDLList);

                                }
                                catch (Exception ex)
                                { }
                            }
                        }
                    }

                    obj.isTure = true;
                    obj.response = listEDDLList;
                    obj.count = count;
                    obj.MachineDispName = machineName.MachineDispName;
                    obj.MachineInvNo = machineName.MachineInvNo;

                }
                catch (Exception ex)
                {
                    obj.isTure = false;
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

        // Getting the shop wise WorkCenter
        public CommonResponse GetShopWiseWorkCenter(int machineId)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<EShopWiseWC> listEShopWiseWC = new List<EShopWiseWC>();
                int CellID = Convert.ToInt32(db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == machineId).Select(x => x.CellId).FirstOrDefault());
                if (CellID != 0)
                {
                    var wcDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == CellID).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
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

        //public CommonResponseWithMachinedesscName GetWo(BatchDetWithMachineName data)
        //{
        //    CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
        //    try
        //    {
        //        List<WoDet> woList = new List<WoDet>();
        //        var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineId && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();
        //        //var wolist = db.Tbllivehmiscreen.Where(m => m.AutoBatchNumber == batchNo && m.IsChecked == 0 && m.IsBatchStart == 0).ToList();
        //        var wolist1 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsBatchStart == 1).FirstOrDefault();
        //        if (wolist1 != null)
        //        {
        //            var wolist = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsChecked == 1).ToList();
        //            if (wolist.Count > 0)
        //            {
        //                foreach (var row in wolist)
        //                {
        //                    int actid = Convert.ToInt32(row.ActivityId);
        //                    int? pid = row.ProcessId;
        //                    var actName = db.Tblactivity.Where(m => m.ActivityId == actid && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
        //                    var pName = db.TblProcess.Where(m => m.ProcessId == pid).Select(m => m.ProcessName).FirstOrDefault();
        //                    WoDet woobj = new WoDet();
        //                    woobj.woNo = row.WorkOrderNo;
        //                    woobj.prevBatchNumber = row.PrevBatchNo;
        //                    woobj.batchNumber = row.AutoBatchNumber;
        //                    woobj.opId = Convert.ToString(row.OperatiorId);
        //                    woobj.ActivityName = actName;
        //                    woobj.ProcessName = pName;
        //                    woobj.BHmiid = row.Bhmiid;
        //                    woobj.operationNo = row.OperationNo;
        //                    woobj.partNo = row.PartNo;
        //                    woobj.project = row.Project;
        //                    woobj.ProdFai = row.ProdFai;
        //                    woobj.DeliverdQty = Convert.ToString(row.DeliveredQty);
        //                    woobj.targetQty = Convert.ToString(row.TargetQty);
        //                    woobj.ProcessQty = Convert.ToString(row.ProcessQty);
        //                    woList.Add(woobj);
        //                }
        //                var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsReworkClicked == 1).FirstOrDefault();
        //                if (getdata != null)
        //                {
        //                    comres.Isrework = true;
        //                }

        //                comres.isTure = true;
        //                comres.response = woList;
        //                comres.processName = (int)wolist[0].ProcessId;
        //                comres.activityName = (int)wolist[0].ActivityId;
        //                comres.MacDispName = machinedet;
        //                comres.Opid = Convert.ToString(wolist[0].OperatiorId);
        //            }
        //        }
        //        else
        //        {
        //            string batchno = data.BatchNo;
        //            //var CreatedWOList = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == batchno && (m.IsChecked == 0 || m.IsChecked== 1)).ToList();
        //            var getBactchHMI = db.Tblbatchhmiscreen.Where(x => (x.IsChecked == 0 || x.IsChecked == 1) && x.AutoBatchNumber == batchno).ToList();
        //            if (getBactchHMI.Count > 0)
        //            {
        //                foreach (var row in getBactchHMI)
        //                {
        //                    int actid = Convert.ToInt32(row.ActivityId);
        //                    int? pid = row.ProcessId;
        //                    var actName = db.Tblactivity.Where(m => m.ActivityId == actid && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
        //                    var pName = db.TblProcess.Where(m => m.ProcessId == pid).Select(m => m.ProcessName).FirstOrDefault();
        //                    WoDet woobj = new WoDet();
        //                    woobj.woNo = row.WorkOrderNo;
        //                    woobj.prevBatchNumber = row.PrevBatchNo;
        //                    woobj.batchNumber = row.AutoBatchNumber;
        //                    woobj.opId = Convert.ToString(row.OperatiorId);
        //                    woobj.ActivityName = actName;
        //                    woobj.ProcessName = pName;
        //                    woobj.BHmiid = row.Bhmiid;
        //                    woobj.operationNo = row.OperationNo;
        //                    woobj.partNo = row.PartNo;
        //                    woobj.project = row.Project;
        //                    woobj.ProdFai = row.ProdFai;
        //                    woobj.DeliverdQty = Convert.ToString(row.DeliveredQty);
        //                    woobj.targetQty = Convert.ToString(row.TargetQty);
        //                    woobj.ProcessQty = Convert.ToString(row.ProcessQty);
        //                    woList.Add(woobj);
        //                }
        //                var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsReworkClicked == 1).FirstOrDefault();
        //                if (getdata != null)
        //                {
        //                    comres.Isrework = true;
        //                }

        //                comres.isTure = true;
        //                comres.response = woList;
        //                string id = Convert.ToString(getBactchHMI[0].ProcessId);
        //                if (id != "")
        //                {
        //                    comres.processName = (int)getBactchHMI[0].ProcessId;
        //                    comres.activityName = (int)getBactchHMI[0].ActivityId;
        //                    comres.MacDispName = machinedet;
        //                    comres.Opid = Convert.ToString(getBactchHMI[0].OperatiorId);
        //                }

        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //} 

        public CommonResponseWithMachinedesscName GetWo(BatchDetWithMachineName data)
        {
            CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
            try
            {
                List<WoDet> woList = new List<WoDet>();
                var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineId && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();

                string batchno = data.BatchNo;
                var getBactchHMI = db.Tblbatchhmiscreen.Where(x => (x.IsChecked == 0 || x.IsChecked == 1) && x.AutoBatchNumber == batchno && x.IsBatchStart == 0).ToList();
                if (getBactchHMI.Count > 0)
                {
                    foreach (var row in getBactchHMI)
                    {
                        int actid = Convert.ToInt32(row.ActivityId);
                        int? pid = row.ProcessId;
                        var actName = db.Tblactivity.Where(m => m.ActivityId == actid && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
                        var pName = db.TblProcess.Where(m => m.ProcessId == pid).Select(m => m.ProcessName).FirstOrDefault();
                        WoDet woobj = new WoDet();
                        woobj.woNo = row.WorkOrderNo;
                        woobj.prevBatchNumber = row.PrevBatchNo;
                        woobj.batchNumber = row.AutoBatchNumber;
                        woobj.opId = Convert.ToString(row.OperatiorId);
                        woobj.ActivityName = actName;
                        woobj.ProcessName = pName;
                        woobj.BHmiid = row.Bhmiid;
                        woobj.operationNo = row.OperationNo;
                        woobj.partNo = row.PartNo;
                        woobj.project = row.Project;
                        woobj.ProdFai = row.ProdFai;
                        woobj.DeliverdQty = Convert.ToString(row.DeliveredQty);
                        woobj.targetQty = Convert.ToString(row.TargetQty);
                        woobj.ProcessQty = Convert.ToString(row.ProcessQty);
                        woobj.pcpNo = row.PcpNo;
                        woobj.RejectionQty = Convert.ToString(row.RejectionQuantity);
                        woList.Add(woobj);
                    }
                    var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsReworkClicked == 1).FirstOrDefault();
                    if (getdata != null)
                    {
                        comres.Isrework = true;
                    }

                    comres.isTure = true;
                    comres.response = woList;
                    string id = Convert.ToString(getBactchHMI[0].ProcessId);
                    if (id != "")
                    {
                        comres.processName = (int)getBactchHMI[0].ProcessId;
                        comres.activityName = (int)getBactchHMI[0].ActivityId;
                        comres.MacDispName = machinedet;
                        comres.Opid = Convert.ToString(getBactchHMI[0].OperatiorId);
                    }

                }
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        //public CommonResponse StartBatch(List<batchdet> data)
        //{
        //    CommonResponse comres = new CommonResponse();
        //    try
        //    {
        //        var shift = GetDateShift();
        //        foreach (var row in data)
        //        {
        //            #region Sequence
        //            ////Tbllivehmiscreen wodet = db.Tbllivehmiscreen.Where(m => m.Hmiid == row.HMIId).FirstOrDefault();
        //            ////if(wodet != null)
        //            ////{
        //            ////    int Hmiid = wodet.Hmiid;
        //            //string OpNo = row.OperationNo;
        //            //string WoNo = row.WoNo;
        //            //string PartNo = row.PartNo;

        //            //int newProcessedQty = 0;
        //            //int PrvProcessQty = 0, PrvDeliveredQty = 0;
        //            //int OperationNoInt = Convert.ToInt32(OpNo);

        //            //var DDLCompletedData = db.Tblddl.Where(m => m.WorkOrder == WoNo && m.MaterialDesc == PartNo && m.OperationNo == OpNo && m.IsCompleted == 1).FirstOrDefault();
        //            //if (DDLCompletedData != null)
        //            //{
        //            //    Tbllivehmiscreen hmiidDataDup = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo && m.MachineId == row.MachineID).FirstOrDefault();
        //            //    comres.isTure = false;
        //            //    comres.response = "Job is Finished for WorkOrder:" + WoNo + " OpNo: " + OpNo + " PartNo:" + PartNo;

        //            //    db.Tbllivehmiscreen.Remove(hmiidDataDup);
        //            //    db.SaveChanges();
        //            //}

        //            ////2017-06-22
        //            //Tbllivehmiscreen HMICompletedData = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo && m.IsWorkInProgress == 1).FirstOrDefault();
        //            //if (HMICompletedData != null)
        //            //{
        //            //    comres.isTure = false;
        //            //    comres.response = "Job is Finished for WorkOrder:" + WoNo + " OpNo: " + OpNo + " PartNo:" + PartNo;

        //            //    db.Tbllivehmiscreen.Remove(HMICompletedData);
        //            //    db.SaveChanges();
        //            //}

        //            //var LiveHMIdata = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo).ToList();
        //            //var HMIdatahistorian = db.Tblhmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo).ToList();

        //            //foreach (var liverow in LiveHMIdata)
        //            //{
        //            //    int InnerOpNo = Convert.ToInt32(liverow.OperationNo);
        //            //    if (OperationNoInt > InnerOpNo)
        //            //    {
        //            //        if (liverow.Date == null) //=> lower OpNo is not submitted.
        //            //        {
        //            //            comres.response = " Submit WONo: " + liverow.WorkOrderNo + " and PartNo: " + liverow.PartNo + " and OperationNo: " + InnerOpNo;
        //            //        }
        //            //    }
        //            //}
        //            //if (HMIdatahistorian.Count > 0)
        //            //{
        //            //    foreach (var historyrow in HMIdatahistorian)
        //            //    {
        //            //        int InnerOpNo = Convert.ToInt32(historyrow.OperationNo);
        //            //        if (OperationNoInt > InnerOpNo)
        //            //        {
        //            //            if (historyrow.Date == null) //=> lower OpNo is not submitted.
        //            //            {
        //            //                comres.response = " Submit WONo: " + historyrow.WorkOrderNo + " and PartNo: " + historyrow.PartNo + " and OperationNo: " + InnerOpNo;
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //            //bool IsInHMI = true;
        //            //var ddllist = db.Tblddl.Where(m => m.WorkOrder == WoNo && m.OperationNo != OpNo && m.MaterialDesc == PartNo && m.IsCompleted == 0).ToList();
        //            //foreach (var ddlrow in ddllist)
        //            //{
        //            //    IsInHMI = true; //reinitialize
        //            //    int InnerOpNo = Convert.ToInt32(ddlrow.OperationNo);
        //            //    if (InnerOpNo < OperationNoInt)
        //            //    {
        //            //        bool IsItWrong = false;
        //            //        var LiveHMIdet = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo).OrderBy(m => m.Hmiid).ToList();
        //            //        var HMIdatahistoriandet = db.Tblhmiscreen.Where(m => m.WorkOrderNo == WoNo && m.PartNo == PartNo && m.OperationNo == OpNo).OrderBy(m => m.Hmiid).ToList();
        //            //        if (LiveHMIdet == null || HMIdatahistoriandet == null)
        //            //        {
        //            //            comres.response = " Select & Start WONo: " + ddlrow.WorkOrder + " and PartNo: " + ddlrow.MaterialDesc + " and OperationNo: " + InnerOpNo;
        //            //        }
        //            //        else
        //            //        {
        //            //            foreach (var rowHMI in LiveHMIdet)
        //            //            {
        //            //                if (rowHMI.Date == null) //=> lower OpNo is not submitted.
        //            //                {
        //            //                    comres.response = " Start WONo: " + ddlrow.WorkOrder + " and PartNo: " + ddlrow.MaterialDesc + " and OperationNo: " + InnerOpNo;
        //            //                }
        //            //            }
        //            //            foreach (var hisrow in HMIdatahistoriandet)
        //            //            {
        //            //                if (hisrow.Date == null)
        //            //                {
        //            //                    comres.response = " Start WONo: " + ddlrow.WorkOrder + " and PartNo: " + ddlrow.MaterialDesc + " and OperationNo: " + InnerOpNo;
        //            //                }
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //            #endregion

        //            #region
        //            //int deliveredQty = 0;
        //            //int TargetQtyNew = Convert.ToInt32(wodet.TargetQty);
        //            //int DeliveredNew = Convert.ToInt32(wodet.DeliveredQty);
        //            //int ProcessedNew = Convert.ToInt32(wodet.ProcessQty);
        //            //newProcessedQty = DeliveredNew + ProcessedNew;
        //            //if (Convert.ToInt32(wodet.IsWorkInProgress) == 1)
        //            //{
        //            //    comres.response = "Job is Finished for WorkOrder:" + WoNo + " OpNo: " + OpNo + " PartNo:" + PartNo;
        //            //    db.Tbllivehmiscreen.Remove(wodet);
        //            //    db.SaveChanges();

        //            //}
        //            //if (TargetQtyNew == newProcessedQty)
        //            //{
        //            //    wodet.TargetQty = newProcessedQty;
        //            //    wodet.ProcessQty = newProcessedQty;
        //            //    wodet.SplitWo = "No";
        //            //    wodet.IsWorkInProgress = 1;
        //            //    wodet.Status = 2;
        //            //    wodet.Time = wodet.Date;
        //            //    wodet.DeliveredQty = 0;

        //            //    db.Entry(wodet).State = EntityState.Modified;
        //            //    db.SaveChanges();


        //            //    //if it existing in DDLList Update 
        //            //    var DDLList = db.Tblddl.Where(m => m.WorkOrder == wodet.WorkOrderNo && m.MaterialDesc == wodet.PartNo && m.OperationNo == wodet.OperationNo && m.IsCompleted == 0).ToList();
        //            //    foreach (var row1 in DDLList)
        //            //    {
        //            //        row1.IsCompleted = 1;
        //            //        db.Entry(row1).State = EntityState.Modified;
        //            //        db.SaveChanges();
        //            //    }
        //            //    comres.response = "Job is Finished for WorkOrder:" + WoNo + " OpNo: " + OpNo + " PartNo:" + PartNo;
        //            //}

        //            //if (TargetQtyNew < newProcessedQty)
        //            //{
        //            //    comres.response = "Previous ProcessedQty :" + newProcessedQty + ". TargetQty Cannot be Less than Processed";
        //            //    wodet.ProcessQty = 0;
        //            //    wodet.Date = null;
        //            //    db.Entry(wodet).State = EntityState.Modified;
        //            //    db.SaveChanges();
        //            //}
        //            #endregion

        //            var batch = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == row.BHmiid).FirstOrDefault();
        //            var batchnodet = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == row.autoBatchNumber).FirstOrDefault();
        //            if (batchnodet.IsActivityFinish == 1)
        //            {
        //                comres.isTure = false;
        //                comres.response = "This Activity is Already Finish so you could not able Start";
        //            }
        //            else
        //            {
        //                if (row.IsChecked == 1)
        //                {
        //                    var sphmidet = db.TblSplivehmiscreen.Where(m => m.MachineId == row.MachineID && m.AutoBatchNumber == row.autoBatchNumber && m.WorkOrderNo == row.WoNo && m.OperationNo == row.OperationNo && m.PartNo == row.PartNo).ToList();
        //                    if (sphmidet.Count == 0)
        //                    {
        //                        if (row.NewAutoBatchNumber != null)
        //                        {
        //                            TblSplivehmiscreen objbat1 = new TblSplivehmiscreen();
        //                            objbat1.ActivityId = Convert.ToInt32(row.ActivityId);
        //                            objbat1.ProcessId = Convert.ToInt32(row.processId);
        //                            objbat1.AutoBatchNumber = row.NewAutoBatchNumber;
        //                            objbat1.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
        //                            objbat1.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        //                            objbat1.Time = null;
        //                            objbat1.WorkOrderNo = row.WoNo;
        //                            objbat1.TargetQty = batch.TargetQty;
        //                            objbat1.Status = 0;
        //                            objbat1.SplitWo = row.SplitWo;
        //                            objbat1.Shift = shift[0];
        //                            objbat1.RejQty = null;
        //                            objbat1.Project = row.project;
        //                            objbat1.ProdFai = row.ProdFai;
        //                            objbat1.ProcessQty = 0;
        //                            objbat1.PestartTime = batch.PestartTime;
        //                            objbat1.PartNo = row.PartNo;
        //                            objbat1.OperationNo = row.OperationNo;
        //                            objbat1.MachineId = batch.MachineId;
        //                            objbat1.IsWorkOrder = 0;
        //                            objbat1.OperatiorId = Convert.ToInt32(row.OperatorId);
        //                            objbat1.TargetQty = Convert.ToInt32(row.TargetQty);
        //                            objbat1.IsBatchStart = 1;
        //                            objbat1.IsUpdate = 1;
        //                            objbat1.IsWorkInProgress = 2;
        //                            objbat1.IsHold = 0;
        //                            objbat1.DoneWithRow = 0;
        //                            objbat1.DeliveredQty = 0;
        //                            objbat1.DdlwokrCentre = batch.DdlwokrCentre;
        //                            db.TblSplivehmiscreen.Add(objbat1);
        //                            db.SaveChanges();

        //                            var macInvNo = db.Tblmachinedetails.Where(m => m.MachineId == row.MachineID).Select(m => m.MachineInvNo).FirstOrDefault();
        //                            var batchdet = db.Tblbatchhmiscreen.Where(m => m.MachineId == row.MachineID && m.AutoBatchNumber == row.NewAutoBatchNumber && m.WorkOrderNo == row.WoNo && m.OperationNo == row.OperationNo && m.PartNo == row.PartNo).FirstOrDefault();
        //                            if (batchdet != null)
        //                            {
        //                                Tblbatchhmiscreen tblbatchhmiscreen = new Tblbatchhmiscreen();
        //                                tblbatchhmiscreen.OperatiorId = 0;
        //                                tblbatchhmiscreen.Status = 0;
        //                                tblbatchhmiscreen.Project = row.project;
        //                                tblbatchhmiscreen.PartNo = row.PartNo;
        //                                tblbatchhmiscreen.OperationNo = row.OperationNo;
        //                                tblbatchhmiscreen.WorkOrderNo = row.WoNo;
        //                                tblbatchhmiscreen.TargetQty = Convert.ToInt32(row.TargetQty);
        //                                tblbatchhmiscreen.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd"); ;
        //                                tblbatchhmiscreen.ProdFai = row.ProdFai;
        //                                tblbatchhmiscreen.PestartTime = DateTime.Now;
        //                                tblbatchhmiscreen.DdlwokrCentre = macInvNo;
        //                                tblbatchhmiscreen.AutoBatchNumber = row.NewAutoBatchNumber;
        //                                tblbatchhmiscreen.MachineId = row.MachineID;
        //                                tblbatchhmiscreen.IsWorkOrder = 0;
        //                                tblbatchhmiscreen.IsChecked = 1;
        //                                db.Tblbatchhmiscreen.Add(tblbatchhmiscreen);
        //                                db.SaveChanges();
        //                            }
        //                            else { }

        //                            batch.IsBatchStart = 1;
        //                            batch.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        //                            batch.ProcessId = row.processId;
        //                            batch.ActivityId = Convert.ToInt32(row.ActivityId);
        //                            db.Entry(batch).State = EntityState.Modified;
        //                            db.SaveChanges();
        //                        }

        //                        TblSplivehmiscreen objbat = new TblSplivehmiscreen();
        //                        objbat.ActivityId = Convert.ToInt32(row.ActivityId);
        //                        objbat.ProcessId = Convert.ToInt32(row.processId);
        //                        objbat.AutoBatchNumber = row.autoBatchNumber;
        //                        objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
        //                        objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        //                        objbat.Time = null;
        //                        objbat.WorkOrderNo = row.WoNo;
        //                        objbat.TargetQty = batch.TargetQty;
        //                        objbat.Status = 0;
        //                        objbat.SplitWo = row.SplitWo;
        //                        objbat.Shift = shift[0];
        //                        objbat.RejQty = null;
        //                        objbat.Project = row.project;
        //                        objbat.ProdFai = row.ProdFai;
        //                        objbat.ProcessQty = 0;
        //                        objbat.PestartTime = batch.PestartTime;
        //                        objbat.PartNo = row.PartNo;
        //                        objbat.OperationNo = row.OperationNo;
        //                        objbat.MachineId = batch.MachineId;
        //                        objbat.IsWorkOrder = 0;
        //                        objbat.OperatiorId = Convert.ToInt32(row.OperatorId);
        //                        objbat.TargetQty = Convert.ToInt32(row.TargetQty);
        //                        objbat.IsBatchStart = 1;
        //                        objbat.IsUpdate = 1;
        //                        objbat.IsWorkInProgress = 2;
        //                        objbat.IsHold = 0;
        //                        objbat.DoneWithRow = 0;
        //                        objbat.DeliveredQty = 0;
        //                        objbat.DdlwokrCentre = batch.DdlwokrCentre;
        //                        db.TblSplivehmiscreen.Add(objbat);
        //                        db.SaveChanges();


        //                        batch.IsBatchStart = 1;
        //                        batch.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        //                        batch.ProcessId = row.processId;
        //                        batch.ActivityId = Convert.ToInt32(row.ActivityId);
        //                        db.Entry(batch).State = EntityState.Modified;
        //                        db.SaveChanges();
        //                    }
        //                    else
        //                    {

        //                    }

        //                }
        //                else
        //                {
        //                    batch.IsChecked = 2;
        //                    db.Entry(batch).State = EntityState.Modified;
        //                    db.SaveChanges();

        //                    var ddl = db.Tblddl.Where(m => m.Bhmiid == row.BHmiid).FirstOrDefault();
        //                    ddl.IsWoselected = 0;
        //                    ddl.Bhmiid = 0;
        //                    db.Entry(ddl).State = EntityState.Modified;
        //                    db.SaveChanges();
        //                }
        //            }

        //            comres.isTure = true;
        //            comres.response = "Batch Started Successfully";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //}

        public CommonResponsewithEror StartBatch(batchdet data)
        {
            CommonResponsewithEror comres = new CommonResponsewithEror();
            try
            {

                bool isAfStart = false;
                bool isPfStart = false;
                int actId = Convert.ToInt32(data.ActivityId);
                //var actId1 = db.Tblactivity.Where(m => m.ActivityName == data.autoBatchNumber && m.Isdeleted == 0).Select(m => m.ActivityId).FirstOrDefault();

                //Checking whether that particular batch is partially finished or not
                var batchno = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.ActivityId == actId && m.IsPartialFinish == 1).FirstOrDefault();
                if (batchno != null)
                {
                    if (batchno.IsPartialFinish == 1)
                    {
                        isPfStart = true;
                    }
                }
                //Checking whether that particular batch is activity finished or not
                var batch1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.ActivityId == actId).ToList();
                if (batch1.Count != 0)
                {
                    foreach (var batchrow in batch1)
                    {
                        if (batchrow.IsActivityFinish == 1)
                        {
                            isAfStart = true;
                        }
                        else
                        {
                            isAfStart = false;
                            break;
                        }
                    }
                }
                var shift = GetDateShift();
                string[] hmiids = data.BHmiid.Split(',');
                foreach (var row in hmiids)
                {
                    int id = Convert.ToInt32(row);
                    var batch = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.Bhmiid == id).FirstOrDefault();
                    if (batch != null)
                    {
                        //foreach (var batchrow in batchhmiids)
                        //{
                        //    int batchid = Convert.ToInt32(batchrow);
                        //    if (id == batchid)
                        //    {

                        //int actId = Convert.ToInt32(data.ActivityId);
                        //var batch1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.ActivityId == actId).FirstOrDefault();
                        //if (batch1 != null)
                        // {
                        //    if (batch1.IsActivityFinish == 1)
                        //    {

                        //    }
                        //    else if (batch.IsPatialFinish == 1)
                        //    {

                        //        TblSplivehmiscreen objbat = new TblSplivehmiscreen();
                        //        objbat.ActivityId = Convert.ToInt32(data.ActivityId);
                        //        objbat.ProcessId = Convert.ToInt32(data.processId);
                        //        objbat.AutoBatchNumber = data.autoBatchNumber;
                        //        objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                        //        objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        //        objbat.Time = null;
                        //        objbat.WorkOrderNo = batch.WorkOrderNo;
                        //        objbat.TargetQty = batch.TargetQty;
                        //        objbat.Status = 0;
                        //        objbat.SplitWo = null;
                        //        objbat.Shift = shift[0];
                        //        objbat.BatchHmiid = batch.Bhmiid;
                        //        objbat.RejQty = null;
                        //        objbat.Project = batch.Project;
                        //        objbat.ProdFai = batch.ProdFai;
                        //        objbat.ProcessQty = 0;
                        //        objbat.PestartTime = batch.PestartTime;
                        //        objbat.PartNo = batch.PartNo;
                        //        objbat.OperationNo = batch.OperationNo;
                        //        objbat.MachineId = batch.MachineId;
                        //        objbat.IsWorkOrder = 0;
                        //        objbat.OperatiorId = data.OperatorId;
                        //        objbat.TargetQty = Convert.ToInt32(batch.TargetQty);
                        //        objbat.IsBatchStart = 1;
                        //        objbat.IsActivityFinish = 0;
                        //        objbat.IsPartialFinish = 0;
                        //        objbat.IsUpdate = 1;
                        //        objbat.IsWorkInProgress = 2;
                        //        objbat.IsHold = 0;
                        //        objbat.DoneWithRow = 0;
                        //        objbat.DeliveredQty = 0;
                        //        objbat.DdlwokrCentre = batch.DdlwokrCentre;
                        //        db.TblSplivehmiscreen.Add(objbat);
                        //        db.SaveChanges();

                        //        batch.IsBatchStart = 1;
                        //        batch.IsChecked = 1;
                        //        db.Entry(batch).State = EntityState.Modified;
                        //        db.SaveChanges();

                        //        comres.isTure = true;
                        //        comres.response = "Batch Created For Partial Finish";
                        //    }
                        //}  

                        //inserting the record when that perticular batch number are partially finished but not activity finish 
                        if (isPfStart == true && isAfStart == false)
                        {
                            int processqty = GetProcessQty(batch.WorkOrderNo, batch.OperationNo, batch.PartNo);
                            TblSplivehmiscreen objbat = new TblSplivehmiscreen();
                            objbat.ActivityId = Convert.ToInt32(data.ActivityId);
                            objbat.ProcessId = Convert.ToInt32(data.processId);
                            objbat.AutoBatchNumber = data.autoBatchNumber;
                            objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                            objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            objbat.Time = null;
                            objbat.WorkOrderNo = batch.WorkOrderNo;
                            objbat.TargetQty = batch.TargetQty;
                            objbat.Status = 0;
                            objbat.SplitWo = null;
                            objbat.Shift = shift[0];
                            objbat.BatchHmiid = batch.Bhmiid;
                            objbat.RejQty = null;
                            objbat.Project = batch.Project;
                            objbat.ProdFai = batch.ProdFai;
                            objbat.PcpNo = batch.PcpNo;
                            if (processqty != 0)
                            {
                                objbat.ProcessQty = processqty;
                            }
                            else
                            {
                                objbat.ProcessQty = 0;
                            }
                            objbat.PestartTime = batch.PestartTime;
                            objbat.PartNo = batch.PartNo;
                            objbat.OperationNo = batch.OperationNo;
                            objbat.MachineId = batch.MachineId;
                            objbat.IsWorkOrder = 0;
                            objbat.OperatiorId = data.OperatorId;
                            objbat.TargetQty = Convert.ToInt32(batch.TargetQty);
                            objbat.IsBatchStart = 1;
                            objbat.IsActivityFinish = 0;
                            objbat.IsPartialFinish = 0;
                            objbat.IsUpdate = 1;
                            objbat.IsWorkInProgress = 2;
                            objbat.IsHold = 0;
                            objbat.DoneWithRow = 0;
                            objbat.DeliveredQty = 0;
                            objbat.DdlwokrCentre = batch.DdlwokrCentre;
                            objbat.PcpNo = batch.PcpNo;
                            objbat.DeliveredQty = 0;
                            objbat.RejectionQuantity = 0;
                            db.TblSplivehmiscreen.Add(objbat);
                            db.SaveChanges();

                            batch.IsBatchStart = 1;
                            batch.IsChecked = 1;
                            db.Entry(batch).State = EntityState.Modified;
                            db.SaveChanges();

                            comres.isTure = true;
                            comres.response = "Activity Created For Partial Finish";

                        }

                        if (isAfStart == true)
                        {
                            comres.isTure = false;
                            comres.errorMsg = "This Activity is Already Finish so you could not able Start";
                        }

                        int actid = Convert.ToInt32(data.ActivityId);
                        var sphmidet = db.TblSplivehmiscreen.Where(m => m.MachineId == data.MachineID && m.AutoBatchNumber == batch.AutoBatchNumber && m.ActivityId == actid && m.WorkOrderNo == batch.WorkOrderNo && m.PartNo == batch.PartNo && m.OperationNo == batch.OperationNo).ToList();
                        if (sphmidet.Count == 0)
                        {
                            if (data.NewAutoBatchNumber != null)
                            {
                                TblSplivehmiscreen objbat1 = new TblSplivehmiscreen();
                                objbat1.ActivityId = Convert.ToInt32(batch.ActivityId);
                                objbat1.ProcessId = Convert.ToInt32(data.processId);
                                objbat1.AutoBatchNumber = data.NewAutoBatchNumber;
                                objbat1.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                                objbat1.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                objbat1.Time = null;
                                objbat1.WorkOrderNo = batch.WorkOrderNo;
                                objbat1.TargetQty = batch.TargetQty;
                                objbat1.Status = 0;
                                objbat1.SplitWo = batch.SplitWo;
                                objbat1.BatchHmiid = batch.Bhmiid;
                                objbat1.Shift = shift[0];
                                objbat1.RejQty = null;
                                objbat1.Project = batch.Project;
                                objbat1.ProdFai = batch.ProdFai;
                                objbat1.ProcessQty = 0;
                                objbat1.PestartTime = batch.PestartTime;
                                objbat1.PartNo = batch.PartNo;
                                objbat1.OperationNo = batch.OperationNo;
                                objbat1.MachineId = data.MachineID;
                                objbat1.IsWorkOrder = 0;
                                objbat1.OperatiorId = data.OperatorId;
                                objbat1.TargetQty = Convert.ToInt32(batch.TargetQty);
                                objbat1.IsBatchStart = 1;
                                objbat1.IsUpdate = 1;
                                objbat1.IsWorkInProgress = 2;
                                objbat1.IsHold = 0;
                                objbat1.DoneWithRow = 0;
                                objbat1.DeliveredQty = 0;
                                objbat1.DdlwokrCentre = batch.DdlwokrCentre;
                                objbat1.PcpNo = batch.PcpNo;
                                db.TblSplivehmiscreen.Add(objbat1);
                                db.SaveChanges();

                                var macInvNo = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineID).Select(m => m.MachineInvNo).FirstOrDefault();
                                var batchdet = db.Tblbatchhmiscreen.Where(m => m.MachineId == data.MachineID && m.AutoBatchNumber == data.NewAutoBatchNumber && m.WorkOrderNo == batch.WorkOrderNo && m.OperationNo == batch.OperationNo && m.PartNo == batch.PartNo).FirstOrDefault();
                                if (batchdet == null)
                                {
                                    Tblbatchhmiscreen tblbatchhmiscreen = new Tblbatchhmiscreen();
                                    tblbatchhmiscreen.OperatiorId = null;
                                    tblbatchhmiscreen.Status = 0;
                                    tblbatchhmiscreen.Project = batch.Project;
                                    tblbatchhmiscreen.PartNo = batch.PartNo;
                                    tblbatchhmiscreen.OperationNo = batch.OperationNo;
                                    tblbatchhmiscreen.WorkOrderNo = batch.WorkOrderNo;
                                    tblbatchhmiscreen.TargetQty = Convert.ToInt32(batch.TargetQty);
                                    tblbatchhmiscreen.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                                    tblbatchhmiscreen.ProdFai = batch.ProdFai;
                                    tblbatchhmiscreen.PestartTime = DateTime.Now;
                                    tblbatchhmiscreen.DdlwokrCentre = macInvNo;
                                    tblbatchhmiscreen.AutoBatchNumber = data.NewAutoBatchNumber;
                                    tblbatchhmiscreen.MachineId = data.MachineID;
                                    tblbatchhmiscreen.IsWorkOrder = 0;
                                    tblbatchhmiscreen.IsChecked = 0;
                                    db.Tblbatchhmiscreen.Add(tblbatchhmiscreen);
                                    db.SaveChanges();
                                }
                                else { }

                                batch.IsBatchStart = 1;
                                batch.IsChecked = 1;
                                batch.OperatiorId = data.OperatorId;
                                batch.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                batch.ProcessId = data.processId;
                                batch.ActivityId = Convert.ToInt32(data.ActivityId);
                                db.Entry(batch).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            TblSplivehmiscreen objbat = new TblSplivehmiscreen();
                            objbat.ActivityId = Convert.ToInt32(data.ActivityId);
                            objbat.ProcessId = Convert.ToInt32(data.processId);
                            objbat.AutoBatchNumber = data.autoBatchNumber;
                            objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                            objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            objbat.Time = null;
                            objbat.WorkOrderNo = batch.WorkOrderNo;
                            objbat.TargetQty = batch.TargetQty;
                            objbat.Status = 0;
                            objbat.SplitWo = null;
                            objbat.Shift = shift[0];
                            objbat.BatchHmiid = batch.Bhmiid;
                            objbat.RejQty = null;
                            objbat.Project = batch.Project;
                            objbat.ProdFai = batch.ProdFai;
                            objbat.ProcessQty = 0;
                            objbat.PestartTime = batch.PestartTime;
                            objbat.PartNo = batch.PartNo;
                            objbat.OperationNo = batch.OperationNo;
                            objbat.MachineId = batch.MachineId;
                            objbat.IsWorkOrder = 0;
                            objbat.OperatiorId = data.OperatorId;
                            objbat.TargetQty = Convert.ToInt32(batch.TargetQty);
                            objbat.IsBatchStart = 1;
                            objbat.IsUpdate = 1;
                            objbat.IsWorkInProgress = 2;
                            objbat.IsHold = 0;
                            objbat.DoneWithRow = 0;
                            objbat.DeliveredQty = 0;
                            objbat.DdlwokrCentre = batch.DdlwokrCentre;
                            db.TblSplivehmiscreen.Add(objbat);
                            db.SaveChanges();


                            batch.IsBatchStart = 1;
                            batch.IsChecked = 1;
                            batch.OperatiorId = data.OperatorId;
                            batch.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            batch.ProcessId = data.processId;
                            batch.ActivityId = Convert.ToInt32(data.ActivityId);
                            db.Entry(batch).State = EntityState.Modified;
                            db.SaveChanges();
                            comres.isTure = true;
                            comres.response = "Activity Started Successfully";
                        }

                        else { }
                        //else
                        //{
                        //    if (batch.IsActivityFinish == 1)
                        //    {
                        //        TblSplivehmiscreen objbat = new TblSplivehmiscreen();
                        //        objbat.ActivityId = Convert.ToInt32(data.ActivityId);
                        //        objbat.ProcessId = Convert.ToInt32(data.processId);
                        //        objbat.AutoBatchNumber = data.autoBatchNumber;
                        //        objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                        //        objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        //        objbat.Time = null;
                        //        objbat.WorkOrderNo = batch.WorkOrderNo;
                        //        objbat.TargetQty = batch.TargetQty;
                        //        objbat.Status = 0;
                        //        objbat.SplitWo = null;
                        //        objbat.Shift = shift[0];
                        //        objbat.BatchHmiid = batch.Bhmiid;
                        //        objbat.RejQty = null;
                        //        objbat.Project = batch.Project;
                        //        objbat.ProdFai = batch.ProdFai;
                        //        objbat.ProcessQty = 0;
                        //        objbat.PestartTime = batch.PestartTime;
                        //        objbat.PartNo = batch.PartNo;
                        //        objbat.OperationNo = batch.OperationNo;
                        //        objbat.MachineId = batch.MachineId;
                        //        objbat.IsWorkOrder = 0;
                        //        objbat.OperatiorId = data.OperatorId;
                        //        objbat.TargetQty = Convert.ToInt32(batch.TargetQty);
                        //        objbat.IsBatchStart = 1;
                        //        objbat.IsActivityFinish = 1;
                        //        objbat.IsUpdate = 1;
                        //        objbat.IsWorkInProgress = 2;
                        //        objbat.IsHold = 0;
                        //        objbat.DoneWithRow = 0;
                        //        objbat.DeliveredQty = 0;
                        //        objbat.DdlwokrCentre = batch.DdlwokrCentre;
                        //        db.TblSplivehmiscreen.Add(objbat);
                        //        db.SaveChanges();

                        //        batch.IsBatchStart = 1;
                        //        batch.IsChecked = 1;
                        //        db.Entry(batch).State = EntityState.Modified;
                        //        db.SaveChanges();

                        //    }
                        //     if (batch.IsPatialFinish == 1)
                        //    {
                        //        TblSplivehmiscreen objbat = new TblSplivehmiscreen();
                        //        objbat.ActivityId = Convert.ToInt32(data.ActivityId);
                        //        objbat.ProcessId = Convert.ToInt32(data.processId);
                        //        objbat.AutoBatchNumber = data.autoBatchNumber;
                        //        objbat.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                        //        objbat.Date = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        //        objbat.Time = null;
                        //        objbat.WorkOrderNo = batch.WorkOrderNo;
                        //        objbat.TargetQty = batch.TargetQty;
                        //        objbat.Status = 0;
                        //        objbat.SplitWo = null;
                        //        objbat.Shift = shift[0];
                        //        objbat.BatchHmiid = batch.Bhmiid;
                        //        objbat.RejQty = null;
                        //        objbat.Project = batch.Project;
                        //        objbat.ProdFai = batch.ProdFai;
                        //        objbat.ProcessQty = 0;
                        //        objbat.PestartTime = batch.PestartTime;
                        //        objbat.PartNo = batch.PartNo;
                        //        objbat.OperationNo = batch.OperationNo;
                        //        objbat.MachineId = batch.MachineId;
                        //        objbat.IsWorkOrder = 0;
                        //        objbat.OperatiorId = data.OperatorId;
                        //        objbat.TargetQty = Convert.ToInt32(batch.TargetQty);
                        //        objbat.IsBatchStart = 1;
                        //        objbat.IsActivityFinish = 0;
                        //        objbat.IsPartialFinish = 1;
                        //        objbat.IsUpdate = 1;
                        //        objbat.IsWorkInProgress = 2;
                        //        objbat.IsHold = 0;
                        //        objbat.DoneWithRow = 0;
                        //        objbat.DeliveredQty = 0;
                        //        objbat.DdlwokrCentre = batch.DdlwokrCentre;
                        //        db.TblSplivehmiscreen.Add(objbat);
                        //        db.SaveChanges();

                        //        batch.IsBatchStart = 1;
                        //        batch.IsChecked = 1;
                        //        db.Entry(batch).State = EntityState.Modified;
                        //        db.SaveChanges();
                        //    }
                        //    comres.isTure = true;
                        //    comres.response = ResourceResponse.SuccessMessage;
                        //    //comres.isTure = false;
                        //    //comres.response = "This Batch is already started";
                        //}
                        //}
                    }
                    // }

                    if (data.unCheckedId != "")
                    {
                        string[] unCheckedHmiids = data.unCheckedId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedHmiids)
                        {
                            int id1 = Convert.ToInt32(uncheckedIdRow);
                            var batchhmiids = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.Bhmiid == id1).FirstOrDefault();
                            if (batchhmiids != null)
                            {
                                int bhmiid = Convert.ToInt32(batchhmiids.Bhmiid);

                                batchhmiids.IsChecked = 2;
                                batchhmiids.IsBatchStart = 0;
                                db.Entry(batchhmiids).State = EntityState.Modified;
                                db.SaveChanges();

                                var sphmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.BatchHmiid == bhmiid).FirstOrDefault();
                                if (sphmi != null)
                                {
                                    sphmi.ActivityId = 0;
                                    sphmi.IsChecked = 2;
                                    sphmi.IsBatchStart = 0;
                                    db.Entry(sphmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                                var ddl = db.Tblddl.Where(m => m.Bhmiid == bhmiid).FirstOrDefault();
                                if (ddl != null)
                                {
                                    ddl.IsWoselected = 0;
                                    ddl.Bhmiid = 0;
                                    db.Entry(ddl).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }

                        }
                    }


                    //var batchhmiids1 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == data.autoBatchNumber && m.IsChecked == 0).ToList();
                    //foreach (var rowbatch in batchhmiids1)
                    //{
                    //    int id = Convert.ToInt32(rowbatch.Bhmiid);
                    //    var ddl = db.Tblddl.Where(m => m.Bhmiid == id).FirstOrDefault();
                    //    if (ddl != null)
                    //    {
                    //        ddl.IsWoselected = 0;
                    //        ddl.Bhmiid = 0;
                    //        db.Entry(ddl).State = EntityState.Modified;
                    //        db.SaveChanges();
                    //    }
                    //}

                    // }

                }
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        public int GetProcessQty(string WoNo, string OpNo, string partNo)
        {
            int processqty = 0;
            var qty = db.TblSplivehmiscreen.Where(m => m.WorkOrderNo == WoNo && m.OperationNo == OpNo && m.PartNo == partNo && m.Status == 1 && m.IsWorkInProgress == 0).OrderByDescending(m => m.Sphmiid).FirstOrDefault();
            if (qty != null)
            {
                int procqty = Convert.ToInt32(qty.ProcessQty);
                int delqty = Convert.ToInt32(qty.DeliveredQty);
                int rejqty = Convert.ToInt32(qty.RejectionQuantity);
                processqty = delqty + procqty + rejqty;
            }
            return processqty;
        }

        //Get Operator ID's based on prefix
        public CommonResponse OperatorDetails(string Operatorid)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<OperatorId> listOperatorId = new List<OperatorId>();
                //List<string> OPDetList = new List<string>();
                var orepatorDet = db.Tbloperatordetails.Where(m => m.IsDeleted == 0 && m.OperatorId.StartsWith(Operatorid)).Select(m => m.OperatorId).ToList().Take(10);
                if (orepatorDet.Count() > 0)
                {
                    foreach (var oprow in orepatorDet)
                    {
                        OperatorId objOperatorId = new OperatorId();
                        objOperatorId.OpId = oprow;
                        listOperatorId.Add(objOperatorId);
                    }
                    comres.isTure = true;
                    comres.response = listOperatorId;
                }
                else
                {
                    comres.isTure = false;
                    comres.response = "No Data Found";
                }
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }


        //public CommonResponseWithMachinedesscName GetBatchNo(int MachineID)
        //{
        //    CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
        //    try
        //    {
        //        var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == MachineID && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();
        //        var shift = GetDateShift();


        //        List<getBatchNo> wolist = new List<getBatchNo>();
        //        //var batchlist = db.TblSplivehmiscreen.Where(m => m.Date != null && m.IsBatchStart == 1 && m.Status != 2 && m.IsWorkInProgress != 1).Select(m => m.AutoBatchNumber).Distinct().ToList();
        //        var batchlist = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID  && m.IsBatchFinish == 0 && m.IsBatchStart == 1 && m.IsChecked != 2).Select(m => m.AutoBatchNumber).Distinct().ToList();
        //        foreach (var row in batchlist)
        //        {
        //            getBatchNo get = new getBatchNo();
        //            get.BatchNo = row;


        //            var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsIdleClicked == 2).FirstOrDefault();
        //            if (getdata != null)
        //            {
        //                comres.IsIdle = true;
        //            }
        //            var getdata1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 2).FirstOrDefault();
        //            if (getdata1 != null)
        //            {
        //                comres.IsPM = true;
        //            }
        //            var getdata2 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 2).FirstOrDefault();
        //            if (getdata2 != null)
        //            {
        //                comres.IsBreakdown = true;
        //            }

        //            var getdata3 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsIdleClicked == 1).FirstOrDefault();
        //            if (getdata3 != null)
        //            {
        //                //getBatchNo get = new getBatchNo();
        //                //get.BatchNo = row;
        //                get.IsIndividualIdle = true;
        //                //wolist.Add(get);
        //            }
        //            var getdata4 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 1).FirstOrDefault();
        //            if (getdata4 != null)
        //            {
        //                //getBatchNo get = new getBatchNo();
        //                //get.BatchNo = row;
        //                get.IsIndividualPM = true;
        //                //wolist.Add(get);
        //            }
        //            var getdata5 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 1).FirstOrDefault();
        //            if (getdata5 != null)
        //            {
        //                //getBatchNo get = new getBatchNo();
        //                //get.BatchNo = row;
        //                get.IsIndividualBreakdown = true;
        //                //wolist.Add(get);
        //            }
        //            var getdata6 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsReworkClicked == 1).FirstOrDefault();
        //            if (getdata6 != null)
        //            {
        //                comres.Isrework = true;
        //            }
        //            var getdata7 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsHoldClicked == 1).FirstOrDefault();
        //            if (getdata7 != null)
        //            {
        //                //getBatchNo get = new getBatchNo();
        //                //get.BatchNo = row;
        //                get.IsIndividualHold = true;
        //                //wolist.Add(get);
        //            }
        //            wolist.Add(get);
        //        }
        //        comres.isTure = true;
        //        comres.response = wolist;
        //        comres.Shift = shift[0];
        //        comres.MacDispName = machinedet;
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //}  

        public CommonResponseWithMachinedesscName GetBatchNo(int MachineID)
        {
            CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
            try
            {
                var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == MachineID && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();
                var shift = GetDateShift();


                List<getBatchNo> wolist = new List<getBatchNo>();
                var batchlist = db.TblSplivehmiscreen.Where(m => m.MachineId == MachineID && m.IsActivityFinish == 0 && m.IsPartialFinish == 0 && m.IsBatchFinish == 0 && m.IsBatchStart == 1&& m.IsChecked !=2).Select(m => m.AutoBatchNumber).Distinct().ToList();
                if (batchlist.Count > 0)
                {
                    foreach (var row in batchlist)
                    {
                        getBatchNo get = new getBatchNo();
                        get.BatchNo = row;


                        var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsIdleClicked == 2).FirstOrDefault();
                        if (getdata != null)
                        {
                            comres.IsIdle = true;
                        }
                        var getdata1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 2).FirstOrDefault();
                        if (getdata1 != null)
                        {
                            comres.IsPM = true;
                        }
                        var getdata2 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 2).FirstOrDefault();
                        if (getdata2 != null)
                        {
                            comres.IsBreakdown = true;
                        }

                        var getdata3 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsIdleClicked == 1).FirstOrDefault();
                        if (getdata3 != null)
                        {
                            //getBatchNo get = new getBatchNo();
                            //get.BatchNo = row;
                            get.IsIndividualIdle = true;
                            //wolist.Add(get);
                        }
                        var getdata4 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 1).FirstOrDefault();
                        if (getdata4 != null)
                        {
                            //getBatchNo get = new getBatchNo();
                            //get.BatchNo = row;
                            get.IsIndividualPM = true;
                            //wolist.Add(get);
                        }
                        var getdata5 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 1).FirstOrDefault();
                        if (getdata5 != null)
                        {
                            //getBatchNo get = new getBatchNo();
                            //get.BatchNo = row;
                            get.IsIndividualBreakdown = true;
                            //wolist.Add(get);
                        }
                        var getdata6 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsReworkClicked == 1).FirstOrDefault();
                        if (getdata6 != null)
                        {
                            comres.Isrework = true;
                        }
                        var getdata7 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsHoldClicked == 1).FirstOrDefault();
                        if (getdata7 != null)
                        {
                            //getBatchNo get = new getBatchNo();
                            //get.BatchNo = row;
                            get.IsIndividualHold = true;
                            //wolist.Add(get);
                        }
                        wolist.Add(get);
                        var woList = db.TblSplivehmiscreen.Where(m => m.Date != null && m.IsBatchStart == 1 && m.AutoBatchNumber == row && m.IsActivityFinish == 0 && m.IsPartialFinish == 0).ToList();
                        foreach (var rolistrow in woList)
                        {
                            if (rolistrow.IsUpdate == 1)
                            {
                                comres.set = true;
                            }
                            else
                            {
                                comres.set = false;
                            }

                        }

                    }
                }

                comres.isTure = true;
                comres.response = wolist;
                comres.Shift = shift[0];
                comres.MacDispName = machinedet;
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        //public CommonResponseWithMachineName GetStartedWo(BatchDetWithMachineName data)
        //{
        //    CommonResponseWithMachineName comres = new CommonResponseWithMachineName();
        //    try
        //    {
        //        List<WoDetails> wolist = new List<WoDetails>();
        //        var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineId && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();

        //        var WOFinishList = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsBatchStart == 1 && (m.IsActivityFinish == 1 || m.IsPartialFinish == 1 && m.ActivityId !=0 )).OrderByDescending(m=>m.ActivityId).ToList();
        //        if (WOFinishList.Count > 0)
        //        {
        //            foreach (var batchrow in WOFinishList)
        //            {
        //                var actName1 = db.Tblactivity.Where(m => m.ActivityId == batchrow.ActivityId && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
        //                var processName1 = db.TblProcess.Where(m => m.ProcessId == batchrow.ProcessId && m.Isdeleted == 0).Select(m => m.ProcessName).FirstOrDefault();
        //                WoDetails woobj1 = new WoDetails();
        //                woobj1.BatchNo = batchrow.AutoBatchNumber;
        //                woobj1.ActivityName = actName1;
        //                woobj1.opId = batchrow.OperatiorId;
        //                woobj1.DeliverdQty = Convert.ToString(batchrow.DeliveredQty);
        //                woobj1.HMIId = batchrow.Sphmiid;
        //                woobj1.operationNo = batchrow.OperationNo;
        //                woobj1.ProcessName = processName1;
        //                woobj1.partNo = batchrow.PartNo;
        //                woobj1.ProcessQty = Convert.ToString(batchrow.ProcessQty);
        //                woobj1.ProdFai = batchrow.ProdFai;
        //                woobj1.project = batchrow.Project;
        //                woobj1.SplitWo = batchrow.SplitWo;
        //                woobj1.targetQty = Convert.ToString(batchrow.TargetQty);
        //                woobj1.woNo = batchrow.WorkOrderNo;
        //                wolist.Add(woobj1);
        //            }

        //        }
        //        else
        //        {
        //            var woList = db.TblSplivehmiscreen.Where(m => m.Date != null && m.IsBatchStart == 1 && m.AutoBatchNumber == data.BatchNo).ToList();
        //            foreach (var row1 in woList)
        //            {
        //                var actName = db.Tblactivity.Where(m => m.ActivityId == row1.ActivityId && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
        //                var processName = db.TblProcess.Where(m => m.ProcessId == row1.ProcessId && m.Isdeleted == 0).Select(m => m.ProcessName).FirstOrDefault();
        //                WoDetails woobj = new WoDetails();
        //                woobj.BatchNo = row1.AutoBatchNumber;
        //                woobj.ActivityName = actName;
        //                woobj.opId = row1.OperatiorId;
        //                woobj.DeliverdQty = Convert.ToString(row1.DeliveredQty);
        //                woobj.HMIId = row1.Sphmiid;
        //                woobj.operationNo = row1.OperationNo;
        //                woobj.ProcessName = processName;
        //                woobj.partNo = row1.PartNo;
        //                woobj.ProcessQty = Convert.ToString(row1.ProcessQty);
        //                woobj.ProdFai = row1.ProdFai;
        //                woobj.project = row1.Project;
        //                woobj.SplitWo = row1.SplitWo;
        //                woobj.targetQty = Convert.ToString(row1.TargetQty);
        //                woobj.woNo = row1.WorkOrderNo;
        //                wolist.Add(woobj);
        //            }
        //        }
        //        var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsIdleClicked == 1).FirstOrDefault();
        //        if (getdata != null)
        //        {
        //            comres.IsIdle = true;
        //        }
        //        var getdata1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IspmClicked == 1).FirstOrDefault();
        //        if (getdata1 != null)
        //        {
        //            comres.IsPM = true;
        //        }
        //        var getdata2 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsBreakdownClicked == 1).FirstOrDefault();
        //        if (getdata2 != null)
        //        {
        //            comres.IsBreakdown = true;
        //        }
        //        var getdata3 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsReworkClicked == 1).FirstOrDefault();
        //        if (getdata3 != null)
        //        {
        //            comres.Isrework = true;
        //        }
        //        var getdata4 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsHoldClicked == 1).FirstOrDefault();
        //        if (getdata4 != null)
        //        {
        //            comres.IsHold = true;
        //        }
        //        comres.isTure = true;
        //        comres.response = wolist;
        //        comres.MacDispName = machinedet;
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //}

        public CommonResponseWithMachineName GetStartedWo(BatchDetWithMachineName data)
        {
            CommonResponseWithMachineName comres = new CommonResponseWithMachineName();
            try
            {
                List<WoDetails> wolist = new List<WoDetails>();
                var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineId && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();

                var woList = db.TblSplivehmiscreen.Where(m => m.Date != null && m.IsBatchStart == 1 && m.AutoBatchNumber == data.BatchNo && m.IsActivityFinish == 0 && m.IsPartialFinish == 0).ToList();
                foreach (var row1 in woList)
                {
                    var actName = db.Tblactivity.Where(m => m.ActivityId == row1.ActivityId && m.Isdeleted == 0).Select(m => m.ActivityName).FirstOrDefault();
                    var processName = db.TblProcess.Where(m => m.ProcessId == row1.ProcessId && m.Isdeleted == 0).Select(m => m.ProcessName).FirstOrDefault();
                    WoDetails woobj = new WoDetails();
                    woobj.BatchNo = row1.AutoBatchNumber;
                    woobj.ActivityName = actName;
                    woobj.opId = row1.OperatiorId;
                    woobj.DeliverdQty = Convert.ToString(row1.DeliveredQty);
                    woobj.HMIId = row1.Sphmiid;
                    woobj.operationNo = row1.OperationNo;
                    woobj.ProcessName = processName;
                    woobj.partNo = row1.PartNo;
                    woobj.ProcessQty = Convert.ToString(row1.ProcessQty);
                    woobj.ProdFai = row1.ProdFai;
                    woobj.project = row1.Project;
                    woobj.SplitWo = row1.SplitWo;
                    woobj.targetQty = Convert.ToString(row1.TargetQty);
                    woobj.woNo = row1.WorkOrderNo;
                    woobj.pcpNo = row1.PcpNo;
                    wolist.Add(woobj);


                }
                var getdata = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsIdleClicked == 1).FirstOrDefault();
                if (getdata != null)
                {
                    comres.IsIdle = true;
                }
                var getdata1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IspmClicked == 1).FirstOrDefault();
                if (getdata1 != null)
                {
                    comres.IsPM = true;
                }
                var getdata2 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsBreakdownClicked == 1).FirstOrDefault();
                if (getdata2 != null)
                {
                    comres.IsBreakdown = true;
                }
                var getdata3 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsReworkClicked == 1).FirstOrDefault();
                if (getdata3 != null)
                {
                    comres.Isrework = true;
                }
                var getdata4 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == data.BatchNo && m.IsHoldClicked == 1).FirstOrDefault();
                if (getdata4 != null)
                {
                    comres.IsHold = true;
                }
                comres.isTure = true;
                comres.response = wolist;
                comres.MacDispName = machinedet;
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }


        #region
        //public CommonResponse IndividualBatchPartial(string Hmiid)
        //{
        //    CommonResponse comres = new CommonResponse();
        //    try
        //    {
        //        int MachineID = 0;
        //        if (IDLEorGenericWorkisON())
        //        {
        //            comres.isTure = false;
        //            comres.response = "Please End IDLE/GenericWork Before Selecting New Work Orders";
        //        }

        //        string[] HMIIDArray = Hmiid.Split(',');
        //        List<string> HMIIDList = Hmiid.Split(',').ToList();
        //        int ExceptionHMIID = 0;
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiiid = Convert.ToInt32(hmiid);
        //            var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiiid).FirstOrDefault();
        //            MachineID = HMIData.MachineId;
        //            if (HMIData.IsHold == 1)
        //            {
        //                comres.isTure = false;
        //                comres.response = "End HOLD before Partial Finish";
        //            }
        //            else if (HMIData.Date != null)
        //            {
        //                int deliveredQty = 0;
        //                if (int.TryParse(Convert.ToString(HMIData.DeliveredQty), out deliveredQty))
        //                {
        //                    int processed = 0;
        //                    int.TryParse(Convert.ToString(HMIData.ProcessQty), out processed);
        //                    if ((deliveredQty + processed) > Convert.ToInt32(HMIData.TargetQty))
        //                    {
        //                        comres.isTure = false;
        //                        comres.response = "DeliveredQty Must be less than Target for " + HMIData.WorkOrderNo;
        //                    }
        //                    else if ((deliveredQty + processed) == Convert.ToInt32(HMIData.TargetQty))
        //                    {
        //                        string wono = HMIData.WorkOrderNo;
        //                        string partno = HMIData.PartNo;
        //                        string opno = HMIData.OperationNo;
        //                        int opnoInt = Convert.ToInt32(HMIData.OperationNo);

        //                        Gen , Seperated String from HMIIDArray
        //                        string hmiids = null;
        //                        for (int hmiidLooper = 0; hmiidLooper < HMIIDArray.Length; hmiidLooper++)
        //                        {
        //                            if (hmiids == null)
        //                            {
        //                                string localhmiidString = HMIIDArray[hmiidLooper];
        //                                if (ExceptionHMIID.ToString() != localhmiidString)
        //                                {
        //                                    hmiids = HMIIDArray[hmiidLooper].ToString();
        //                                }
        //                            }
        //                            else
        //                            {
        //                                if (ExceptionHMIID.ToString() != HMIIDArray[hmiidLooper])
        //                                {
        //                                    hmiids += "," + HMIIDArray[hmiidLooper].ToString();
        //                                }
        //                            }
        //                        }
        //                        var ddldet = db.Tblddl.Where(m => m.WorkOrder == wono && m.OperationNo == opno && m.MaterialDesc == partno && m.IsCompleted == 0).OrderBy(m => m.OperationNo).ToList();
        //                        foreach (var row in ddldet)
        //                        {
        //                            int InnerOpNo = Convert.ToInt32(row.OperationNo);
        //                            string InnerOpNoString = Convert.ToString(row.OperationNo);
        //                            bool IsInHMI = true;
        //                            if (opnoInt > InnerOpNo)
        //                            {
        //                                var hmidet = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == wono && m.OperationNo == opno && m.PartNo == partno && m.IsWorkInProgress == 1).FirstOrDefault();
        //                                if (hmidet != null)
        //                                {
        //                                    IsInHMI = true;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "Enter Delivered Quantity";
        //                }
        //            }
        //            else if (HMIData.Date == null)
        //            {// Do Nothing. Just to Skip our Extra Empty Row
        //                if (ExceptionHMIID == 0)
        //                {
        //                    ExceptionHMIID = Convert.ToInt32(HMIData.Hmiid);
        //                }
        //                else
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "Please Start All WorkOrders Before PartialFinish";
        //                }
        //            }
        //            else
        //            {
        //                comres.isTure = false;
        //                comres.response = "Please Start All WorkOrders Before PartialFinish";
        //            }
        //        }
        //        Check if there are any row to partialFinish(Empty Screen)
        //        if (HMIIDList.Count == 1)
        //        {
        //            foreach (var hmiid in HMIIDArray)
        //            {
        //                int hmiidi = Convert.ToInt32(hmiid);
        //                if (hmiidi == ExceptionHMIID)
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "There are no WorkOrder to Finish";
        //                }
        //            }
        //        }
        //        bool IsSequenceOK = true;
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiidi = Convert.ToInt32(hmiid);
        //            {
        //                var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiidi).FirstOrDefault();
        //                int DelivQty = 0, TargQty = 0, ProcessedQty = 0;

        //                int.TryParse(HMIData.TargetQty.ToString(), out TargQty);
        //                int.TryParse(HMIData.DeliveredQty.ToString(), out DelivQty);
        //                int.TryParse(HMIData.ProcessQty.ToString(), out ProcessedQty);

        //                if (TargQty >= (DelivQty + ProcessedQty))
        //                {
        //                    string WoNo = HMIData.WorkOrderNo;
        //                    string partNo = HMIData.PartNo;
        //                    int OpNo = Convert.ToInt32(HMIData.OperationNo);
        //                    foreach (var hmiid1 in HMIIDArray)
        //                    {
        //                        int hmiidiInner = Convert.ToInt32(hmiid1);
        //                        if (hmiidi != ExceptionHMIID)
        //                        {
        //                            var HMIDataInner = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiidiInner && m.WorkOrderNo == WoNo && m.PartNo == partNo && m.Hmiid != hmiidi).FirstOrDefault();
        //                            if (HMIDataInner != null)
        //                            {
        //                                int opNoInner = Convert.ToInt32(HMIDataInner.OperationNo);
        //                                if (opNoInner < OpNo)
        //                                {
        //                                    int DelivQty1 = 0, TargQty1 = 0, ProcessedQty1 = 0;
        //                                    int.TryParse(HMIDataInner.TargetQty.ToString(), out TargQty1);
        //                                    int.TryParse(HMIDataInner.DeliveredQty.ToString(), out DelivQty1);
        //                                    int.TryParse(HMIDataInner.ProcessQty.ToString(), out ProcessedQty1);
        //                                    if (TargQty1 < (DelivQty + ProcessedQty))
        //                                    {
        //                                        comres.isTure = false;
        //                                        comres.response = "Click JobFinish, All WorkOrders have (Delivered + Processed) = Target.";
        //                                        IsSequenceOK = false;
        //                                        break;
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                    if (!IsSequenceOK)
        //                    {
        //                        break;
        //                    }
        //                }
        //                else
        //                {
        //                    IsSequenceOK = false;
        //                    break;
        //                }
        //            }
        //        }
        //        bool AllREligableForJF = true;
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiidi = Convert.ToInt32(hmiid);
        //            if (hmiidi != ExceptionHMIID)
        //            {
        //                var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiidi).FirstOrDefault();
        //                int DelivQty = 0;
        //                int TargQty = 0;
        //                int ProcessedQty = 0;

        //                int.TryParse(HMIData.TargetQty.ToString(), out TargQty);
        //                int.TryParse(HMIData.DeliveredQty.ToString(), out DelivQty);
        //                int.TryParse(HMIData.ProcessQty.ToString(), out ProcessedQty);

        //                if (TargQty == (DelivQty + ProcessedQty))
        //                {
        //                }
        //                else
        //                {
        //                    AllREligableForJF = false;
        //                    break;
        //                }
        //            }
        //        }
        //        if (AllREligableForJF)
        //        {
        //            comres.isTure = false;
        //            comres.response = "Click JobFinish, All WorkOrders have (Delivered + Processed) = Target.";
        //        }

        //        List<string> MacHierarchy = GetMacHierarchyData(MachineID);

        //        Update ProcessedQty if WorkOrder Available in Different WorkCenter
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiidi = Convert.ToInt32(hmiid);
        //            if (hmiidi != ExceptionHMIID)
        //            {
        //                var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiidi).FirstOrDefault();
        //                int DelivQty = 0, TargQty = 0, ProcessedQty = 0;
        //                int.TryParse(HMIData.DeliveredQty.ToString(), out DelivQty);
        //                int HMIMacID = HMIData.MachineId;

        //                #region If Its as SingleWO
        //                var SimilarWOData = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == HMIData.WorkOrderNo && m.OperationNo == HMIData.OperationNo && m.PartNo == HMIData.PartNo && m.MachineId != MachineID && m.IsWorkInProgress == 2).ToList();
        //                foreach (var row in SimilarWOData)
        //                {
        //                    int InnerProcessed = row.ProcessQty;
        //                    int FinalProcessed = DelivQty + InnerProcessed;
        //                    if (FinalProcessed < row.TargetQty)
        //                    {
        //                        if (row.IsWorkInProgress == 2)
        //                        {
        //                            int ProcessQty = FinalProcessed;
        //                            int hmid = row.Hmiid;
        //                            db.Entry(row).State = EntityState.Modified;
        //                            dbsimilar.SaveChanges();
        //                        }
        //                    }
        //                    else
        //                    {
        //                        comres.isTure = false;
        //                        comres.response = " Same WorkOrder in Machine: " + MacHierarchy[3] + "->" + MacHierarchy[4] + " has ProcessedQty :" + InnerProcessed;
        //                    }
        //                }
        //                #endregion


        //            }
        //        }
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiidi = Convert.ToInt32(hmiid);
        //            if (hmiidi != ExceptionHMIID)
        //            {
        //                var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiidi).FirstOrDefault();
        //                int DelivQty = 0;
        //                int TargQty = 0;
        //                int ProcessedQty = 0;

        //                int.TryParse(HMIData.TargetQty.ToString(), out TargQty);
        //                int.TryParse(HMIData.DeliveredQty.ToString(), out DelivQty);
        //                int.TryParse(HMIData.ProcessQty.ToString(), out ProcessedQty);

        //                if (TargQty == (DelivQty + ProcessedQty))
        //                {
        //                    string wono = HMIData.WorkOrderNo;
        //                    string partno = HMIData.PartNo;
        //                    string opno = HMIData.OperationNo;
        //                    int opnoInt = Convert.ToInt32(HMIData.OperationNo);

        //                    HMIData.Status = 2;
        //                    HMIData.IsWorkInProgress = 1;
        //                    var ddldata = db.Tblddl.Where(m => m.WorkOrder == wono && m.MaterialDesc == partno && m.OperationNo == opno).FirstOrDefault();
        //                    if (ddldata != null)
        //                    {
        //                        int ddlID = ddldata.Ddlid;
        //                        ddldata.IsCompleted = 1;
        //                        db.Entry(ddldata).State = EntityState.Modified;
        //                        db.SaveChanges();
        //                    }
        //                }
        //                else if (TargQty > (DelivQty + ProcessedQty))
        //                {
        //                    HMIData.Status = 1;
        //                    HMIData.IsWorkInProgress = 0;
        //                }
        //                else
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "Delivered + Processed Cannot be Greater than Target for :" + HMIData.WorkOrderNo;
        //                }
        //                HMIData.Time = DateTime.Now;
        //                int HMIID = HMIData.Hmiid;
        //                db.Entry(HMIData).State = EntityState.Modified;
        //                db.SaveChanges();
        //                comres.isTure = true;
        //                comres.response = "Job is Partially finished Successfully";
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return comres;
        //}

        //public CommonResponse IndividualBatchFinish(string Hmiid)
        //{
        //    CommonResponse comres = new CommonResponse();
        //    int MachineId = 1;
        //    try
        //    {
        //        if (IDLEorGenericWorkisON())
        //        {
        //            comres.isTure = false;
        //            comres.response = "Please End IDLE/GenericWork Before Selecting New Work Orders";
        //        }
        //        string[] HMIIDArray = Hmiid.Split(',');

        //        int ExceptionHMIID = 0;
        //        foreach (var hmiid in HMIIDArray)
        //        {
        //            int hmiid1 = Convert.ToInt32(hmiid);
        //            var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiid1).FirstOrDefault();
        //            if (HMIData.IsHold == 1)
        //            {
        //                comres.isTure = false;
        //                comres.response = "End HOLD before Clicking Job Finish";
        //            }
        //            else if (HMIData.Date != null)
        //            {
        //                int deliveredQty = 0, processQty = 0;
        //                int.TryParse(Convert.ToString(HMIData.ProcessQty), out processQty);
        //                if (int.TryParse(Convert.ToString(HMIData.DeliveredQty), out deliveredQty))
        //                {
        //                    if ((deliveredQty + processQty) == Convert.ToInt32(HMIData.TargetQty))
        //                    {
        //                    }
        //                    else
        //                    {
        //                        comres.isTure = false;
        //                        comres.response = "DeliveredQty Must be equal to Target Qty";
        //                    }
        //                }
        //                else
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "Enter Delivered Quantity";
        //                }
        //            }
        //            else if (HMIData.Date == null)
        //            {// Do Nothing. Just to Skip our Extra Empty Row
        //                if (ExceptionHMIID == 0)
        //                {
        //                    ExceptionHMIID = Convert.ToInt32(HMIData.Hmiid);
        //                }
        //                else
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "Please enter all Details Before StartAll is Clicked.";
        //                }
        //            }
        //            else
        //            {
        //                comres.isTure = false;
        //                comres.response = "Please Start All WorkOrders Before JobFinish";
        //            }
        //        }
        //        Check if its  Empty Screen and JobFinish is clicked.
        //        if (HMIIDArray.Length == 1)
        //        {
        //            foreach (var hmiid in HMIIDArray)
        //            {
        //                int hmiidi = Convert.ToInt32(hmiid);
        //                if (hmiidi == ExceptionHMIID)
        //                {
        //                    comres.isTure = false;
        //                    comres.response = "There are no WorkOrder to Finish";
        //                }
        //            }
        //        }
        //        string hmiids = null;
        //        for (int hmiid = 0; hmiid < HMIIDArray.Length; hmiid++)
        //        {
        //            if (hmiids == null)
        //            {
        //                string localhmiidString = HMIIDArray[hmiid];
        //                if (ExceptionHMIID.ToString() != localhmiidString)
        //                {
        //                    hmiids = HMIIDArray[hmiid].ToString();
        //                }
        //            }
        //            else
        //            {
        //                if (ExceptionHMIID.ToString() != HMIIDArray[hmiid])
        //                {
        //                    hmiids += "," + HMIIDArray[hmiid].ToString();
        //                }
        //            }
        //        }
        //        Get hmiids(as comma seperated string) in ascending order based on wono,partno,Opno.
        //       hmiids = GetOrderedHMIIDs(hmiids);

        //        List<Tbllivehmiscreen> Hmiobj = new List<Tbllivehmiscreen>();
        //        foreach (var row in hmiids)
        //        {
        //            var hmidet = db.Tbllivehmiscreen.Where(m => m.Hmiid == row).OrderBy(m => new { m.WorkOrderNo, m.PartNo, m.OperationNo }).FirstOrDefault();
        //            Hmiobj.Add(hmidet);
        //        }
        //        foreach (var rowOuter in Hmiobj)
        //        {
        //            int HMIId = rowOuter.Hmiid;
        //            var SimilarWOData = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == rowOuter.WorkOrderNo && m.OperationNo == rowOuter.OperationNo && m.PartNo == rowOuter.PartNo && m.MachineId != MachineId && m.IsWorkInProgress == 2).FirstOrDefault();

        //            if (SimilarWOData != null)
        //            {
        //                int InnerMacID = Convert.ToInt32(db.Tbllivehmiscreen.Where(m => m.Hmiid == SimilarWOData.Hmiid).Select(m => m.MachineId).FirstOrDefault());
        //                var MacDispName = Convert.ToString(db.Tblmachinedetails.Where(m => m.MachineId == InnerMacID).Select(m => m.MachineDispName).FirstOrDefault());
        //                comres.isTure = false;
        //                comres.response = " Same WorkOrder in Machine: " + MacDispName + " , So you cannot JobFinish ";
        //            }
        //        }
        //        foreach (var rowOuter in Hmiobj)
        //        {
        //            int hmiid1 = Convert.ToInt32(rowOuter.Hmiid);
        //            string woNo = Convert.ToString(rowOuter.WorkOrderNo);
        //            string opNo = Convert.ToString(rowOuter.OperationNo);
        //            string partNo = Convert.ToString(rowOuter.PartNo);

        //            Logic to check sequence of JF Based on WONo, PartNo and OpNo.
        //            int OperationNoInt = Convert.ToInt32(opNo);

        //            var WIP1 = db.Tblddl.Where(m => m.WorkOrder == woNo && m.MaterialDesc == partNo && m.OperationNo != opNo && m.IsCompleted == 0).OrderBy(m => m.OperationNo).ToList();
        //            foreach (var row in WIP1)
        //            {
        //                int InnerOpNo = Convert.ToInt32(row.OperationNo);
        //                string ddlopno = row.OperationNo;
        //                string InnerOpNoString = Convert.ToString(row.OperationNo);
        //                if (hmiids.Contains(InnerOpNoString))
        //                { }
        //                else
        //                {
        //                    if (OperationNoInt > InnerOpNo)
        //                    {
        //                        int PrvProcessQty = 0, PrvDeliveredQty = 0, TotalProcessQty = 0, ishold = 0;
        //                        #region new code
        //                        here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
        //                        int isHMIFirst = 2; //default NO History for that wo,pn,on

        //                        var hmiData = db.Tbllivehmiscreen.Where(m => m.WorkOrderNo == woNo && m.PartNo == partNo && m.OperationNo == ddlopno && m.IsWorkInProgress != 2).OrderByDescending(m => m.Time).Take(1).ToList();

        //                        if (hmiData.Count > 0) // now check for greatest amongst
        //                        {
        //                            DateTime hmiDateTime = Convert.ToDateTime(hmiData[0].Time);
        //                            isHMIFirst = 0;
        //                        }

        //                        else if (hmiData.Count > 0)
        //                        {
        //                            isHMIFirst = 0;
        //                        }

        //                        else if (isHMIFirst == 0)
        //                        {
        //                            string delivString = Convert.ToString(hmiData[0].DeliveredQty);
        //                            int delivInt = 0;
        //                            int.TryParse(delivString, out delivInt);

        //                            string processString = Convert.ToString(hmiData[0].ProcessQty);
        //                            int procInt = 0;
        //                            int.TryParse(processString, out procInt);

        //                            PrvProcessQty += procInt;
        //                            PrvDeliveredQty += delivInt;

        //                            ishold = hmiData[0].IsHold;
        //                            ishold = ishold == 2 ? 0 : ishold;
        //                        }
        //                        else
        //                        {
        //                            no previous delivered or processed qty so Do Nothing.
        //                        }
        //                        #endregion
        //                        TotalProcessQty = PrvProcessQty + PrvDeliveredQty;
        //                        var hmiPFed = db.tblhmiscreens.Where(m => m.Work_Order_No == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.Time).FirstOrDefault();

        //                        if (Convert.ToInt32(row.TargetQty) == TotalProcessQty)
        //                        {
        //                            #region
        //                            if (isHMIFirst == 0 && Convert.ToInt32(row.TargetQty) < Convert.ToInt32(hmiData[0].TargetQty))
        //                            {
        //                                Tbllivehmiscreen tblh = new Tbllivehmiscreen();

        //                                tblh.CorrectedDate = row.CorrectedDate;
        //                                tblh.Date = DateTime.Now;
        //                                tblh.Time = DateTime.Now;
        //                                tblh.PestartTime = DateTime.Now;
        //                                tblh.DdlwokrCentre = hmiData[0].DdlwokrCentre;
        //                                tblh.DeliveredQty = 0;
        //                                tblh.DoneWithRow = 1;
        //                                tblh.IsHold = 0;
        //                                tblh.IsMultiWo = hmiData[0].IsMultiWo;
        //                                tblh.IsUpdate = 1;
        //                                tblh.IsWorkInProgress = 1;
        //                                tblh.IsWorkOrder = hmiData[0].IsWorkOrder;
        //                                tblh.MachineId = hmiData[0].MachineId;
        //                                tblh.OperationNo = hmiData[0].OperationNo;
        //                                tblh.OperatiorId = hmiData[0].OperatiorId;
        //                                tblh.OperatorDet = hmiData[0].OperatorDet;
        //                                tblh.PartNo = hmiData[0].PartNo;
        //                                tblh.ProcessQty = TotalProcessQty;
        //                                tblh.ProdFai = hmiData[0].ProdFai;
        //                                tblh.Project = hmiData[0].Project;
        //                                tblh.RejQty = hmiData[0].RejQty;
        //                                tblh.Shift = hmiData[0].Shift;
        //                                tblh.SplitWo = hmiData[0].SplitWo;
        //                                tblh.Status = hmiData[0].Status;
        //                                tblh.TargetQty = TotalProcessQty;
        //                                tblh.WorkOrderNo = hmiData[0].WorkOrderNo;
        //                                db.Tbllivehmiscreen.Add(tblh);
        //                                db.SaveChanges();
        //                            }
        //                            #endregion
        //                            int ddlID = row.Ddlid;
        //                            row.IsCompleted = 1;
        //                            db.Entry(row).State = EntityState.Modified;
        //                            db.SaveChanges();
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        foreach (var rowOuter in Hmiobj)
        //        {
        //            int hmiid1 = Convert.ToInt32(rowOuter.Hmiid);
        //            #region 2017-02-07

        //            string woNo = Convert.ToString(rowOuter.WorkOrderNo);
        //            string opNo = Convert.ToString(rowOuter.OperationNo);
        //            string partNo = Convert.ToString(rowOuter.PartNo);

        //            Logic to check sequence of JF Based on WONo, PartNo and OpNo.
        //            int OperationNoInt = Convert.ToInt32(opNo);

        //            var WIP1 = db.Tblddl.Where(m => m.WorkOrder == woNo && m.MaterialDesc == partNo && m.OperationNo != opNo && m.IsCompleted == 0).OrderBy(m => m.OperationNo).ToList();
        //            foreach (var row in WIP1)
        //            {
        //                int InnerOpNo = Convert.ToInt32(row.OperationNo);
        //                string InnerOpNoString = Convert.ToString(row.OperationNo);
        //                if (hmiids.Contains(InnerOpNoString))
        //                { }
        //                else
        //                {
        //                    if (OperationNoInt > InnerOpNo)
        //                    {
        //                        if (row.IsCompleted != 1) //=> lower OpNo is not Finished.
        //                        {
        //                            comres.isTure = false;
        //                            comres.response = " Finish WONo: " + row.WorkOrder + " and PartNo: " + row.MaterialDesc + " and OperationNo: " + InnerOpNo;
        //                        }
        //                        else { }
        //                    }
        //                }
        //            }
        //            #endregion

        //            var HMIData = db.Tbllivehmiscreen.Where(m => m.Hmiid == hmiid1).FirstOrDefault();
        //            string wono = HMIData.WorkOrderNo;
        //            string partno = HMIData.PartNo;
        //            string opno = HMIData.OperationNo;
        //            var ddldata = db.Tblddl.Where(m => m.WorkOrder == wono && m.MaterialDesc == partno && m.OperationNo == opno).FirstOrDefault();
        //            if (ddldata != null)
        //            {
        //                int ddlID = ddldata.Ddlid;
        //                ddldata.IsCompleted = 1;
        //                db.Entry(ddldata).State = EntityState.Modified;
        //                db.SaveChanges();
        //            }

        //            HMIData.Status = 2;
        //            HMIData.IsWorkInProgress = 1;
        //            HMIData.SplitWo = "No";
        //            HMIData.Time = DateTime.Now;
        //            db.Entry(HMIData).State = EntityState.Modified;
        //            db.SaveChanges();
        //            comres.isTure = true;
        //            comres.response = "Job Finished successfully";
        //        }

        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return comres;
        //}
        #endregion

        public CommonResponsewithEror IndividualPartialFinish(ActivityFinish AFData)
        {
            CommonResponsewithEror comres = new CommonResponsewithEror();
            try
            {

                //bool ispartial = false;
                //var actId = db.Tblactivity.Where(m => m.ActivityName == AFData.activityName && m.Isdeleted == 0).Select(m => m.ActivityId).FirstOrDefault();
                //var batch1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.ActivityId == actId).FirstOrDefault();
                //if (batch1 != null)
                //{
                //    if (batch1.IsPartialFinish == 1)
                //    {

                //        ispartial = true;
                //    }
                //}
                if (IDLEorGenericWorkisON())
                {
                    comres.isTure = false;
                    comres.response = "Please End IDLE/GenericWork Before Selecting New Work Orders";
                }
                string[] hmiids = AFData.HMIId.Split(',');
                List<string> hmiidlist = hmiids.ToList();
                foreach (var row in hmiids)
                {
                    int id = Convert.ToInt32(row);
                    var batchhmiids = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Sphmiid == id).FirstOrDefault();
                    if (batchhmiids != null)
                    {
                        var HMIData = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).FirstOrDefault();
                        if (HMIData.IsHold == 1)
                        {
                            comres.isTure = false;
                            comres.response = "End HOLD before Clicking Job Finish";
                        }
                        int total = (int)batchhmiids.DeliveredQty + batchhmiids.ProcessQty;
                        int targetQty = (int)batchhmiids.TargetQty;
                        if (total == targetQty)
                        {
                            HMIData.ActivityEndTime = DateTime.Now;
                            HMIData.Time = DateTime.Now;
                            HMIData.IsActivityFinish = 1;
                            db.Entry(HMIData).State = EntityState.Modified;
                            db.SaveChanges();

                            var hmidat1 = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                            if (hmidat1 != null)
                            {
                                hmidat1.IsActivityFinish = 1;
                                hmidat1.IsBatchStart = 0;
                                db.Entry(hmidat1).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            bool ret = GetActivityForActivity(HMIData);

                            if (ret == true)
                            {
                                HMIData.Time = DateTime.Now;
                                HMIData.IsWorkInProgress = 1;
                                HMIData.Status = 2;
                                HMIData.IsBatchFinish = 1;
                                db.Entry(HMIData).State = EntityState.Modified;
                                db.SaveChanges();

                                var hmidatails = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                                if (hmidatails != null)
                                {
                                    hmidatails.IsBatchFinish = 1;
                                    hmidatails.IsBatchStart = 1;
                                    db.Entry(hmidatails).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                        }
                        //else if (total == targetQty)
                        //{
                        //    HMIData.ActivityEndTime = DateTime.Now;
                        //    HMIData.Time = DateTime.Now;
                        //    HMIData.IsActivityFinish = 1;
                        //    db.Entry(HMIData).State = EntityState.Modified;
                        //    db.SaveChanges();

                        //    var hmidat1 = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                        //    if (hmidat1 != null)
                        //    {
                        //        hmidat1.IsActivityFinish = 1;
                        //        hmidat1.IsBatchStart = 0;
                        //        db.Entry(hmidat1).State = EntityState.Modified;
                        //        db.SaveChanges();
                        //    }
                        //}
                        else if (total < targetQty)
                        {
                            HMIData.Time = DateTime.Now;
                            HMIData.IsWorkInProgress = 0;
                            HMIData.Status = 1;
                            HMIData.IsPartialFinish = 1;
                            db.Entry(HMIData).State = EntityState.Modified;
                            db.SaveChanges();

                            var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).Select(m => m.AutoBatchNumber).FirstOrDefault();
                            var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                            if (hmi != null)
                            {
                                hmi.IsReworkClicked = 0;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            var hmidat = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                            if (hmidat != null)
                            {
                                hmidat.IsPatialFinish = 1;
                                hmidat.IsBatchStart = 0;
                                db.Entry(hmidat).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                        }

                        // }
                    }

                    if (AFData.unCheckedId != "")
                    {
                        string[] unCheckedHmiids = AFData.unCheckedId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedHmiids)
                        {
                            int id1 = Convert.ToInt32(uncheckedIdRow);
                            var batchhmiidsdet = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Bhmiid == id1).FirstOrDefault();
                            if (batchhmiids != null)
                            {
                                int bhmiid = Convert.ToInt32(batchhmiidsdet.Bhmiid);

                                batchhmiidsdet.IsChecked = 2;
                                batchhmiidsdet.IsPatialFinish = 0;
                                db.Entry(batchhmiids).State = EntityState.Modified;
                                db.SaveChanges();

                                var sphmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.BatchHmiid == bhmiid).FirstOrDefault();
                                if (sphmi != null)
                                {
                                    sphmi.IsPartialFinish = 0;
                                    sphmi.IsChecked = 2;
                                    db.Entry(sphmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                                var ddl = db.Tblddl.Where(m => m.Bhmiid == bhmiid).FirstOrDefault();
                                if (ddl != null)
                                {
                                    ddl.IsWoselected = 0;
                                    ddl.Bhmiid = 0;
                                    db.Entry(ddl).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                }

                //    var batchhmiids1 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.IsChecked == 0).ToList();
                //foreach (var rowbatch in batchhmiids1)
                //{
                //    int rowid = Convert.ToInt32(rowbatch.Bhmiid);
                //    var batch = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == rowid).FirstOrDefault();
                //    if (batch != null)
                //    {
                //        batch.IsPatialFinish = 0;
                //        db.Entry(batch).State = EntityState.Modified;
                //        db.SaveChanges();
                //    }


                //    var ddl = db.Tblddl.Where(m => m.Bhmiid == rowid).FirstOrDefault();
                //    if (ddl != null)
                //    {
                //        ddl.IsWoselected = 0;
                //        ddl.Bhmiid = 0;
                //        db.Entry(ddl).State = EntityState.Modified;
                //        db.SaveChanges();
                //    }

                //}

                // }


                comres.isTure = true;
                comres.response = "item Updated Successfully";
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comres.isTure = false;
            }
            return comres;
        }

        //List<string> GetMacHierarchyData(int MachineID)
        //{
        //    List<string> HierarchyData = new List<string>();
        //    //1st get PlantName or -
        //    //2nd get ShopName or -
        //    //3rd get CellName or -
        //    //4th get MachineName.

        //    var machineData = db.Tblmachinedetails.Where(m => m.MachineId == MachineID).FirstOrDefault();
        //    int PlantID = Convert.ToInt32(machineData.PlantId);
        //    string name = "-";
        //    name = db.Tblplant.Where(m => m.PlantId == PlantID).Select(m => m.PlantName).FirstOrDefault();
        //    HierarchyData.Add(name);
        //    int shopid = Convert.ToInt32(machineData.PlantId);
        //    //string ShopIDString = Convert.ToString(machineData.ShopID);
        //    //    int value;
        //    //    if (int.TryParse(ShopIDString, out value))
        //    //    {
        //    name = db.Tblshop.Where(m => m.ShopId == shopid).Select(m => m.ShopName).FirstOrDefault();
        //    HierarchyData.Add(name.ToString());
        //    //}
        //    // else
        //    // {
        //    //     HierarchyData.Add("-");
        //    // }
        //    int CellID = Convert.ToInt32(machineData.PlantId);
        //    //string CellIDString = Convert.ToString(machineData.CellID);
        //    //    if (int.TryParse(CellIDString, out value))
        //    //    {
        //    name = db.Tblcell.Where(m => m.CellId == CellID).Select(m => m.CellName).FirstOrDefault();
        //    HierarchyData.Add(name.ToString());
        //    //}
        //    //else
        //    //{
        //    //    HierarchyData.Add("-");
        //    //}
        //    HierarchyData.Add(Convert.ToString(machineData.MachineInvNo));
        //    HierarchyData.Add(Convert.ToString(machineData.MachineDispName));
        //    return HierarchyData;
        //}

        public bool IDLEorGenericWorkisON()
        {

            //Check if the Machine is in IDLE or GenericWork , We shouldn't allow them to do other activities if these are ON.
            //bool ItsOn = false;
            int MachineID = 1;
            var GWToView = db.TblSpgenericworkentry.Where(m => m.MachineId == MachineID).OrderByDescending(m => m.SpgwentryId).FirstOrDefault();
            if (GWToView != null) //implies genericwork is running
            {
                if (GWToView.DoneWithRow == 0)
                {
                    //ItsOn = true;
                    return true;
                }
            }
            var IdleToView = db.TblSplivelossofentry.Where(m => m.MachineId == MachineID).OrderByDescending(m => m.SplossId).FirstOrDefault();
            if (IdleToView != null) //implies idle is running
            {
                if (IdleToView.DoneWithRow == 0)
                {
                    //ItsOn = true;
                    return true;
                }
            }

            return false;
        }

        //string GetOrderedHMIIDs(string hmiids)
        //{
        //    string retHMIIDs = null;
        //    if (hmiids != null)
        //    {
        //        List<Tbllivehmiscreen> obj = new List<Tbllivehmiscreen>();
        //        foreach (var row in hmiids)
        //        {
        //            var hmidet = db.Tbllivehmiscreen.Where(m => m.Hmiid == row).OrderBy(m => new { m.WorkOrderNo, m.PartNo, m.OperationNo }).FirstOrDefault();
        //            obj.Add(hmidet);
        //        }
        //        for (int id = 0; id < obj.Count; id++)
        //        {
        //            if (retHMIIDs == null)
        //            {
        //                retHMIIDs = Convert.ToString(obj[id].Hmiid);
        //            }
        //            else
        //            {
        //                retHMIIDs += "," + Convert.ToString(obj[id].Hmiid);
        //            }
        //        }
        //    }

        //    return retHMIIDs;
        //}

        public CommonResponsewithEror IndividualBatchFinish(ActivityFinish AFData)
        {
            CommonResponsewithEror comres = new CommonResponsewithEror();
            try
            {

                string[] hmiids = AFData.HMIId.Split(',');
                bool bf = false;
                foreach (var hmiidrow in hmiids)
                {
                    int id = Convert.ToInt32(hmiidrow);
                    var batchhmiids = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Sphmiid == id).FirstOrDefault();
                    if (batchhmiids != null)
                    {
                        int total = (int)batchhmiids.DeliveredQty + batchhmiids.ProcessQty;
                        int targetQty = (int)batchhmiids.TargetQty;
                        if (total == targetQty)
                        {
                            bf = true;
                        }
                        else
                        {
                            bf = false;
                            break;
                        }
                    }
                }
                if (bf == true)
                {
                    foreach (var row in hmiids)
                    {
                        int id = Convert.ToInt32(row);
                        var batchhmiids = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Sphmiid == id).FirstOrDefault();
                        if (batchhmiids != null)
                        {
                            var HMIData = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).FirstOrDefault();
                            if (HMIData.IsHold == 1)
                            {
                                comres.isTure = false;
                                comres.response = "End HOLD before Clicking Job Finish";
                            }
                            else if (HMIData.Date != null)
                            {
                                int deliveredQty = 0;
                                if (int.TryParse(Convert.ToString(HMIData.DeliveredQty), out deliveredQty))
                                {
                                    int processed = 0;
                                    int.TryParse(Convert.ToString(HMIData.ProcessQty), out processed);
                                    if ((deliveredQty + processed) <= Convert.ToInt32(HMIData.TargetQty))
                                    {
                                        comres.isTure = false;
                                        comres.response = "DeliveredQty Must be equal";
                                    }
                                }
                            }
                            HMIData.Time = DateTime.Now;
                            HMIData.IsWorkInProgress = 1;
                            HMIData.Status = 2;
                            HMIData.IsBatchFinish = 1;
                            HMIData.IsActivityFinish = 1;
                            db.Entry(HMIData).State = EntityState.Modified;
                            db.SaveChanges();

                            var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).Select(m => m.AutoBatchNumber).FirstOrDefault();
                            var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                            if (hmi != null)
                            {
                                hmi.IsReworkClicked = 0;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            var hmidat = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                            if (hmidat != null)
                            {
                                hmidat.IsBatchFinish = 1;
                                hmidat.IsActivityFinish = 1;      //when they do batch Finish Activity End Time will not update
                                hmidat.IsBatchStart = 1;
                                db.Entry(hmidat).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            var ddl = db.Tblddl.Where(m => m.WorkOrder == HMIData.WorkOrderNo && m.OperationNo == HMIData.OperationNo && m.MaterialDesc == HMIData.PartNo && m.IsCompleted == 0).FirstOrDefault();
                            if (ddl != null)
                            {
                                ddl.IsCompleted = 1;
                                db.Entry(ddl).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            var tblmodedata = db.TblSplivemodedb.Where(m => m.IsDeleted == 0 && m.IsCompleted == 0 && m.BatchNumber == AFData.autoBatchNumber).OrderByDescending(m => m.StartTime).ToList();
                            if (tblmodedata.Count != 0)
                            {
                                foreach (var row1 in tblmodedata)
                                {
                                    row1.EndTime = DateTime.Now;
                                    row1.IsCompleted = 1;
                                    double diff = DateTime.Now.Subtract(Convert.ToDateTime(row1.StartTime)).TotalSeconds;
                                    row1.DurationInSec = (int)diff;
                                    db.Entry(row1).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                    if (AFData.unCheckedId != "")
                    {
                        string[] unCheckedHmiids = AFData.unCheckedId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedHmiids)
                        {
                            int id = Convert.ToInt32(uncheckedIdRow);
                            var batchhmiids = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Bhmiid == id).FirstOrDefault();
                            if (batchhmiids != null)
                            {
                                int bhmiid = Convert.ToInt32(batchhmiids.Bhmiid);

                                batchhmiids.IsChecked = 2;
                                batchhmiids.IsBatchFinish = 0;
                                db.Entry(batchhmiids).State = EntityState.Modified;
                                db.SaveChanges();

                                var sphmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.BatchHmiid == bhmiid).FirstOrDefault();
                                if (sphmi != null)
                                {
                                    sphmi.Status = 0;
                                    sphmi.IsWorkInProgress = 2;
                                    sphmi.IsChecked = 2;
                                    sphmi.IsBatchFinish = 0;
                                    db.Entry(sphmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                                var ddl = db.Tblddl.Where(m => m.Bhmiid == bhmiid).FirstOrDefault();
                                if (ddl != null)
                                {
                                    ddl.IsWoselected = 0;
                                    ddl.Bhmiid = 0;
                                    db.Entry(ddl).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                    comres.isTure = true;
                    comres.response = "item Updated Successfully";
                }
                else if (bf == false)
                {
                    comres.isTure = false;
                    comres.errorMsg = "Delivered qty + Process qty must be equal to Total Qty";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comres.isTure = false;
            }
            return comres;
        }

        public CommonResponsewithEror IndividualActivityFinish(ActivityFinish AFData)
        {
            CommonResponsewithEror comres = new CommonResponsewithEror();
            try
            {
                bool act = false;
                string[] hmiids = AFData.HMIId.Split(',');
                foreach (var hmiidrow in hmiids)
                {
                    int id = Convert.ToInt32(hmiidrow);
                    var batchhmiids = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Sphmiid == id).FirstOrDefault();
                    if (batchhmiids != null)
                    {
                        int total = Convert.ToInt32(batchhmiids.DeliveredQty + batchhmiids.ProcessQty);
                        int targetQty = Convert.ToInt32(batchhmiids.TargetQty);

                        if (total == targetQty)
                        {
                            act = true;  // For Activity Finish and Batch Finish
                        }
                        else
                        {
                            act = false;
                            break;
                        }
                    }
                }
                if (act == true)
                {
                    foreach (var row in hmiids)
                    {

                        int id = Convert.ToInt32(row);
                        var batchhmiids = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Sphmiid == id).FirstOrDefault();
                        if (batchhmiids != null)
                        {
                            var HMIData = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).FirstOrDefault();
                            if (HMIData != null)
                            {
                                if (HMIData.IsHold == 1)
                                {
                                    comres.isTure = false;
                                    comres.response = "End HOLD before Clicking Job Finish";
                                }
                            }

                            var actdet = db.Tblactivity.Where(m => m.ActivityId == HMIData.ActivityId).Select(m => m.OptionalAct).FirstOrDefault();
                            if (actdet != null)
                            {
                                if (actdet == "No")
                                {
                                    HMIData.ActivityEndTime = DateTime.Now;
                                    HMIData.IsActivityFinish = 1;
                                    HMIData.Time = DateTime.Now;
                                    db.Entry(HMIData).State = EntityState.Modified;
                                    db.SaveChanges();

                                    bool ret = GetActivityForActivity(HMIData);
                                    log.Error("Activity is true:" + ret);
                                    if (ret == true)
                                    {
                                        //Doing the Batch Finish as well as Activity Finish for that particular Activity
                                        var sphmidet = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.IsChecked !=2).ToList();
                                        foreach(var batchrow in sphmidet)
                                        {
                                            batchrow.Time = DateTime.Now;
                                            batchrow.IsWorkInProgress = 1;
                                            batchrow.Status = 2;
                                            batchrow.IsBatchFinish = 1;
                                            db.Entry(batchrow).State = EntityState.Modified;
                                            db.SaveChanges();
                                        }
                                       

                                        var hmidat1 = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                                        if (hmidat1 != null)
                                        {
                                            hmidat1.IsBatchFinish = 1;
                                            hmidat1.IsBatchStart = 1;
                                            hmidat1.IsActivityFinish = 1;
                                            db.Entry(hmidat1).State = EntityState.Modified;
                                            db.SaveChanges();
                                        }

                                        var ddldet = db.Tblddl.Where(m => m.WorkOrder == hmidat1.WorkOrderNo && m.OperationNo == hmidat1.OperationNo).FirstOrDefault();
                                        if(ddldet != null)
                                        {
                                            ddldet.IsCompleted = 1;
                                            db.SaveChanges();
                                        }

                                        var tblmodedata = db.TblSplivemodedb.Where(m => m.IsDeleted == 0 && m.IsCompleted == 0 && m.BatchNumber == AFData.autoBatchNumber).OrderByDescending(m => m.StartTime).ToList();
                                        if (tblmodedata.Count != 0)
                                        {
                                            foreach (var row1 in tblmodedata)
                                            {
                                                row1.EndTime = DateTime.Now;
                                                row1.IsCompleted = 1;
                                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(row1.StartTime)).TotalSeconds;
                                                row1.DurationInSec = (int)diff;
                                                db.Entry(row1).State = EntityState.Modified;
                                                db.SaveChanges();
                                            }
                                        }
                                        comres.isTure = true;
                                        comres.response = "item Updated Successfully";
                                    }
                                    else
                                    {
                                        HMIData.ActivityEndTime = DateTime.Now;
                                        HMIData.IsActivityFinish = 1;
                                        HMIData.Time = DateTime.Now;
                                        db.Entry(HMIData).State = EntityState.Modified;
                                        db.SaveChanges();

                                        comres.isTure = true;
                                        comres.response = "item Updated Successfully";
                                        var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).Select(m => m.AutoBatchNumber).FirstOrDefault();
                                        var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                                        if (hmi != null)
                                        {
                                            hmi.IsReworkClicked = 0;
                                            db.Entry(hmi).State = EntityState.Modified;
                                            db.SaveChanges();
                                        }
                                        var hmidat = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                                        if (hmidat != null)
                                        {
                                            hmidat.IsActivityFinish = 1;
                                            hmidat.IsBatchStart = 0;
                                            db.Entry(hmidat).State = EntityState.Modified;
                                            db.SaveChanges();
                                        }
                                    }
                                }
                                else if (actdet == "Yes")
                                {
                                    HMIData.ActivityEndTime = DateTime.Now;
                                    HMIData.IsActivityFinish = 1;
                                    HMIData.Time = DateTime.Now;
                                    db.Entry(HMIData).State = EntityState.Modified;
                                    db.SaveChanges();

                                    comres.isTure = true;
                                    comres.response = "item Updated Successfully";
                                    var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).Select(m => m.AutoBatchNumber).FirstOrDefault();
                                    var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                                    if (hmi != null)
                                    {
                                        hmi.IsReworkClicked = 0;
                                        db.Entry(hmi).State = EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                    var hmidat = db.Tblbatchhmiscreen.Where(m => m.Bhmiid == HMIData.BatchHmiid).FirstOrDefault();
                                    if (hmidat != null)
                                    {
                                        hmidat.IsActivityFinish = 1;
                                        hmidat.IsBatchStart = 0;
                                        db.Entry(hmidat).State = EntityState.Modified;
                                        db.SaveChanges();
                                    }
                                }
                            }
                        }
                    }

                    if (AFData.unCheckedId != "")
                    {
                        string[] unCheckedHmiids = AFData.unCheckedId.Split(',');
                        foreach (var uncheckedIdRow in unCheckedHmiids)
                        {
                            int id = Convert.ToInt32(uncheckedIdRow);
                            var batchhmiids = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.Bhmiid == id).FirstOrDefault();
                            if (batchhmiids != null)
                            {
                                int bhmiid = Convert.ToInt32(batchhmiids.Bhmiid);

                                batchhmiids.IsChecked = 2;
                                batchhmiids.IsActivityFinish = 0;
                                db.Entry(batchhmiids).State = EntityState.Modified;
                                db.SaveChanges();

                                var sphmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == AFData.autoBatchNumber && m.BatchHmiid == bhmiid).FirstOrDefault();
                                if (sphmi != null)
                                {
                                    sphmi.ActivityId = 0;
                                    sphmi.IsChecked = 2;
                                    sphmi.IsActivityFinish = 0;
                                    db.Entry(sphmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                                var ddl = db.Tblddl.Where(m => m.Bhmiid == bhmiid).FirstOrDefault();
                                if (ddl != null)
                                {
                                    ddl.IsWoselected = 0;
                                    ddl.Bhmiid = 0;
                                    db.Entry(ddl).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                        }

                    }
                }
                else if (act == false)
                {
                    comres.isTure = false;
                    comres.errorMsg = "Delivered qty + Process qty must be equal to Total Qty";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comres.isTure = false;
            }
            return comres;
        }


        public bool GetActivityForActivity(TblSplivehmiscreen data)
        {
            log.Error("Checking for Activity:");
            bool check = false;
            List<int> actdet = new List<int>();

            var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == data.MachineId).FirstOrDefault();

            //Get all the mandatory Activities for that particular machine
            actdet = db.Tblactivity.Where(m => m.MachineId == data.MachineId && m.OptionalAct == "No").Select(m => m.ActivityId).ToList();
            if (actdet.Count == 0)
            {
                log.Error("actdet:" + actdet.Count);
                //Get all the mandatory Activities for that particular Cell
                actdet = db.Tblactivity.Where(m => m.CellId == machinedet.CellId && m.OptionalAct == "No").Select(m => m.ActivityId).ToList();
                if (actdet.Count == 0)
                {
                    log.Error("actdet:" + actdet.Count);
                    //Get all the mandatory Activities for that particular Shop
                    actdet = db.Tblactivity.Where(m => m.ShopId == machinedet.ShopId && m.OptionalAct == "No").Select(m => m.ActivityId).ToList();

                }
            }
            if (actdet.Count >= 0)
            {
                int i = 1;  //for testing purpose
                foreach (var actrow in actdet)
                {
                    //Checking whether all the mandatory Activities are completed or not in tblsplivehmiscreen table
                    var sphmi = db.TblSplivehmiscreen.Where(m => m.Sphmiid == data.Sphmiid && m.ActivityId == actrow).FirstOrDefault();
                    if (sphmi != null)
                    {
                        if (sphmi.IsActivityFinish == 1)
                        {
                            log.Error("sphmi table true:" + i);
                            i++;
                            check = true;
                        }
                        else
                        {
                            log.Error("sphmi table false:" + i);
                            i++;
                            check = false;
                            break;
                        }
                    }
                    else
                    {
                        check = false;
                    }

                }
            }
            log.Error("return" + check);
            return check;
        }

        public CommonResponse SelectReWorkorder(string hmiids)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] Hmiarry = hmiids.Split(',');
                List<string> HMIIDList = hmiids.Split(',').ToList();
                foreach (var row in HMIIDList)
                {
                    int id = Convert.ToInt32(row);
                    var hmidet = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).FirstOrDefault();
                    hmidet.IsWorkOrder = 1;
                    db.Entry(hmidet).State = EntityState.Modified;
                    db.SaveChanges();

                    var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == id).Select(m => m.AutoBatchNumber).FirstOrDefault();
                    var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                    hmi.IsReworkClicked = 1;
                    db.Entry(hmi).State = EntityState.Modified;
                    db.SaveChanges();
                }
                obj.isTure = true;
                obj.response = "Rework is Started";
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponsewithEror GetPlannedMaintaince(int MachineID)
        {
            CommonResponsewithEror objcom = new CommonResponsewithEror();
            try
            {
                bool plan = false;
                var breakdownclick = db.TblSplivehmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 1 && m.IsBatchFinish == 0 && m.IsPartialFinish == 0 && m.IsActivityFinish == 0).ToList();
                if (breakdownclick.Count != 0)
                {
                    foreach (var breakdownrow in breakdownclick)
                    {
                        if (breakdownrow.IsIdleClicked == 1 || breakdownrow.IsIdleClicked == 2)
                        {
                            objcom.isTure = false;
                            objcom.errorMsg = "Idle is already Running so you cannot able to start Planned Maintainance, first End Idle then start Planned Maintainance";
                            plan = true;
                        }
                        else if (breakdownrow.IsBreakdownClicked == 1 || breakdownrow.IsBreakdownClicked == 2)
                        {
                            objcom.isTure = false;
                            objcom.errorMsg = "BreakDown is already Running so you cannot able to start Planned Maintainance, first End BreakDown then start Planned Maintainance";
                            plan = true;
                        }
                        else if (breakdownrow.IsHoldClicked == 1)
                        {
                            objcom.isTure = false;
                            objcom.errorMsg = "Hold is already Running so you cannot able to start Planned Maintainance, first End Hold then start Planned Maintainance";
                            plan = true;
                        }
                        else if (breakdownrow.IspmClicked == 2)
                        {
                            objcom.isTure = false;
                            objcom.errorMsg = "Overall Planned Maintainance is already Running so you cannot able to start Indivual Planned Maintainance, first End Overall Planned Maintainance then start Individual Planned Maintainance";
                            plan = true;
                        }
                    }

                }
                if (plan == true)
                {

                }
                else
                {
                    string CorrectedDate = null;
                    Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                    TimeSpan Start = StartTime.StartTime;
                    if (Start <= DateTime.Now.TimeOfDay)
                    {
                        CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                    }
                    var curMode = db.TblSpbreakdown.Where(m => m.MachineId == MachineID).Where(m => m.CorrectedDate == CorrectedDate && m.EndTime == null).OrderByDescending(m => m.SpbreakdownId);
                    {

                    }
                    int currentId = 0;

                    foreach (var j in curMode)
                    {
                        currentId = j.SpbreakdownId;
                        string mode = j.MessageCode;
                        if (mode != "PM")
                        {
                            currentId = j.SpbreakdownId;
                            DateTime time = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            Tblbreakdown tbd = db.Tblbreakdown.Find(currentId);
                            tbd.EndTime = time;
                            db.Entry(tbd).State = EntityState.Modified;
                            db.SaveChanges();
                            break;
                        }
                        //else if (mode == "PM")
                        //{
                        //    Session["ModeError"] = "Machine is in Planned Maintenance Mode";
                        //    return RedirectToAction("Index");
                        //}

                    }
                    var brkdown = db.TblSpbreakdown.Where(m => m.MachineId == MachineID).Where(m => m.EndTime == null && m.MessageCode == "PM");
                    if (brkdown.Count() != 0)
                    {
                        int brekdnID = 0;
                        foreach (var a in brkdown)
                        {
                            brekdnID = a.SpbreakdownId;
                        }
                        TblSpbreakdown brekdn = db.TblSpbreakdown.Find(brekdnID);
                        //CheckLastOneHourDownTime(id);

                        List<breakdoen> losslist1 = new List<breakdoen>();
                        var losscode1 = db.Tbllossescodes.Where(m => m.IsDeleted == 0).Where(m => m.MessageType == "PM" && m.LossCodeId == brekdn.BreakDownCode).ToList();
                        foreach (var row in losscode1)
                        {
                            breakdoen obj = new breakdoen();
                            obj.Losscodeid = row.LossCodeId;
                            obj.Losscode = row.LossCodeDesc;
                            losslist1.Add(obj);
                        }
                        if (losslist1 != null)
                        {
                            objcom.isTure = true;
                            objcom.response = losslist1;
                        }
                    }
                    else
                    {
                    }
                    //CheckLastOneHourDownTime(id);

                    List<breakdoen> losslist = new List<breakdoen>();
                    var losscode = db.Tbllossescodes.Where(m => m.IsDeleted == 0).Where(m => m.MessageType == "PM").ToList();
                    foreach (var row in losscode)
                    {
                        breakdoen obj = new breakdoen();
                        obj.Losscodeid = row.LossCodeId;
                        obj.Losscode = row.LossCodeDesc;
                        losslist.Add(obj);
                    }
                    if (losslist != null)
                    {
                        objcom.isTure = true;
                        objcom.response = losslist;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                objcom.isTure = false;
            }
            return objcom;
        }

        public CommonResponsewithEror Plannedmaintain(PM data)
        {
            CommonResponsewithEror comobj = new CommonResponsewithEror();
            try
            {
                int breakdownid = 0;
                string CorrectedDate = null;
                var shift = GetDateShift();
                Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan Start = StartTime.StartTime;
                if (Start <= DateTime.Now.TimeOfDay)
                {
                    CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                string[] hmiids = data.BatchNo.Split(',');
                foreach (var row in hmiids)
                {
                    if (data.Status == "Start")
                    {
                        string CorrectedDate1 = CorrectedDate;
                        var LossData = db.Tbllossescodes.Where(m => m.LossCodeId == data.LossCodeId).FirstOrDefault();
                        var tblbreakdown = db.TblSpbreakdown.Where(m => m.BatchNo == row && m.DoneWithRow == 0).FirstOrDefault();
                        if (tblbreakdown == null)
                        {
                            TblSpbreakdown lossentry = new TblSpbreakdown();
                            lossentry.Shift = shift[0];
                            lossentry.BatchNo = row;
                            lossentry.MessageCode = (LossData.LossCode).ToString();
                            lossentry.BreakDownCode = 120;
                            lossentry.DoneWithRow = 0;
                            lossentry.MessageDesc = "PM";
                            lossentry.MachineId = data.MachineID;
                            lossentry.StartTime = DateTime.Now;
                            lossentry.CorrectedDate = CorrectedDate1;
                            db.TblSpbreakdown.Add(lossentry);
                            db.SaveChanges();
                            breakdownid = lossentry.SpbreakdownId;

                            //update the endtime for the last mode of this machine 
                            var tblmodedata = db.TblSplivemodedb.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).ToList();
                            foreach (var row1 in tblmodedata)
                            {
                                row1.EndTime = DateTime.Now;
                                row1.IsCompleted = 1;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(row1.StartTime)).TotalSeconds;
                                row1.DurationInSec = (int)diff;
                                db.Entry(row1).State = EntityState.Modified;
                                db.SaveChanges();


                            }
                            //Code to save this event to tblmode table
                            TblSplivemodedb tm = new TblSplivemodedb();
                            tm.MachineId = data.MachineID;
                            tm.BatchNumber = row;
                            var userid = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tm.CorrectedDate = CorrectedDate;
                            tm.ColorCode = "red";
                            tm.InsertedOn = DateTime.Now;
                            tm.InsertedBy = userid.UserId;
                            tm.IsPm = 1;
                            tm.IsCompleted = 0;
                            tm.IsDeleted = 0;
                            tm.Mode = "BREAKDOWN";
                            tm.StartTime = DateTime.Now;
                            db.TblSplivemodedb.Add(tm);
                            db.SaveChanges();

                            if (data.IndividualClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                                hmi.IspmClicked = 1;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            else if (data.OverallClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                                hmi.IspmClicked = 2;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            comobj.isTure = true;
                            comobj.response = "Red";

                        }
                        else { }
                    }
                    else
                    {

                        //breakdownid = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineID && m.BreakDownCode == data.LossCodeId).OrderByDescending(m => m.StartTime).Select(m => m.SpbreakdownId).FirstOrDefault();
                        var hmiPm = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 1).FirstOrDefault();
                        if (hmiPm != null)
                        {
                            breakdownid = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineID && m.DoneWithRow == 0 && m.BatchNo == row).OrderByDescending(m => m.StartTime).Select(m => m.SpbreakdownId).FirstOrDefault();
                            if (breakdownid != 0)
                            {
                                var breakdowndet = db.TblSpbreakdown.Where(m => m.SpbreakdownId == breakdownid).FirstOrDefault();
                                if (breakdowndet != null)
                                {
                                    breakdowndet.CorrectedDate = CorrectedDate;
                                    breakdowndet.EndTime = DateTime.Now;
                                    breakdowndet.DoneWithRow = 1;
                                    breakdowndet.MachineId = data.MachineID;
                                    db.Entry(breakdowndet).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                            var tblmodedata1 = db.TblSplivemodedb.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).ToList();
                            if (tblmodedata1.Count > 0)
                            {
                                foreach (var row1 in tblmodedata1)
                                {
                                    row1.EndTime = DateTime.Now;
                                    row1.IsCompleted = 1;
                                    double diff = DateTime.Now.Subtract(Convert.ToDateTime(row1.StartTime)).TotalSeconds;
                                    row1.DurationInSec = (int)diff;
                                    db.Entry(row1).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                            TblSplivemodedb tmIDLE1 = new TblSplivemodedb();
                            var userid1 = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tmIDLE1.ColorCode = "Green";
                            tmIDLE1.BatchNumber = row;
                            tmIDLE1.InsertedOn = DateTime.Now;
                            tmIDLE1.InsertedBy = userid1.UserId;
                            tmIDLE1.StartTime = DateTime.Now;
                            tmIDLE1.CorrectedDate = CorrectedDate;
                            tmIDLE1.IsCompleted = 0;
                            tmIDLE1.IsDeleted = 0;
                            tmIDLE1.MachineId = data.MachineID;
                            tmIDLE1.Mode = "PowerOn";
                            db.TblSplivemodedb.Add(tmIDLE1);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IspmClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Individual PM has not been started";
                                }
                            }

                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && (m.IspmClicked == 2 || m.IspmClicked == 1)).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IspmClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Overall PM has not been started";
                                }
                            }
                            comobj.isTure = true;
                            comobj.response = "Blue";
                        }
                        else
                        {
                            breakdownid = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineID && m.DoneWithRow == 0 && m.BatchNo == row).OrderByDescending(m => m.StartTime).Select(m => m.SpbreakdownId).FirstOrDefault();
                            if (breakdownid != 0)
                            {
                                var breakdowndet = db.TblSpbreakdown.Where(m => m.SpbreakdownId == breakdownid).FirstOrDefault();
                                if (breakdowndet != null)
                                {
                                    breakdowndet.CorrectedDate = CorrectedDate;
                                    breakdowndet.EndTime = DateTime.Now;
                                    breakdowndet.DoneWithRow = 1;
                                    breakdowndet.MachineId = data.MachineID;
                                    db.Entry(breakdowndet).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }

                            var tblmodedata = db.TblSplivemodedb.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).ToList();
                            if (tblmodedata.Count > 0)
                            {
                                foreach (var row1 in tblmodedata)
                                {
                                    row1.EndTime = DateTime.Now;
                                    row1.IsCompleted = 1;
                                    double diff = DateTime.Now.Subtract(Convert.ToDateTime(row1.StartTime)).TotalSeconds;
                                    row1.DurationInSec = (int)diff;
                                    db.Entry(row1).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                            }
                            TblSplivemodedb tmIDLE = new TblSplivemodedb();
                            var userid = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tmIDLE.ColorCode = "Green";
                            tmIDLE.BatchNumber = row;
                            tmIDLE.InsertedOn = DateTime.Now;
                            tmIDLE.InsertedBy = userid.UserId;
                            tmIDLE.StartTime = DateTime.Now;
                            tmIDLE.CorrectedDate = CorrectedDate;
                            tmIDLE.IsCompleted = 0;
                            tmIDLE.IsDeleted = 0;
                            tmIDLE.MachineId = data.MachineID;
                            tmIDLE.Mode = "PowerOn";
                            db.TblSplivemodedb.Add(tmIDLE);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IspmClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Individual PM has not been started";
                                }
                            }

                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IspmClicked == 2).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IspmClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Overall PM has not been started";
                                }
                            }
                            comobj.isTure = true;
                            comobj.response = "Blue";
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comobj.isTure = false;
            }
            return comobj;
        }

        public CommonResponsewithEror BreaksdownEntry(PM data)
        {
            CommonResponsewithEror comobj = new CommonResponsewithEror();
            try
            {
                int breakdownid = 0;
                string CorrectedDate = null;
                var shift = GetDateShift();
                Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan Start = StartTime.StartTime;
                if (Start <= DateTime.Now.TimeOfDay)
                {
                    CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                string[] hmiids = data.BatchNo.Split(',');
                foreach (var row in hmiids)
                {
                    if (data.Status == "Start")
                    {
                        var tblbreakdown = db.TblSpbreakdown.Where(m => m.BatchNo == row && m.DoneWithRow == 0).FirstOrDefault();
                        if (tblbreakdown == null)
                        {
                            var breakdata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeId == data.LossCodeId).FirstOrDefault();
                            string msgcode = breakdata.LossCode;
                            string msgdesc = breakdata.LossCodeDesc;

                            TblSpbreakdown tb = new TblSpbreakdown();
                            tb.BreakDownCode = data.LossCodeId;
                            tb.BatchNo = row;
                            tb.CorrectedDate = CorrectedDate;
                            tb.DoneWithRow = 0;
                            tb.MachineId = data.MachineID;
                            tb.MessageCode = msgcode;
                            tb.MessageDesc = msgdesc;
                            tb.Shift = shift[0];
                            tb.StartTime = DateTime.Now;
                            db.TblSpbreakdown.Add(tb);
                            db.SaveChanges();

                            //Code to End PreviousMode(Production Here) & save this event to tblmode table
                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.IsCompleted = 1;
                                modedata.EndTime = DateTime.Now;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.Entry(modedata).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            TblSplivemodedb tm = new TblSplivemodedb();
                            tm.MachineId = data.MachineID;
                            tm.CorrectedDate = CorrectedDate;
                            tm.BatchNumber = row;
                            var userid = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tm.InsertedBy = userid.UserId;
                            tm.StartTime = DateTime.Now;
                            tm.ColorCode = "red";
                            tm.IsBreakdown = 1;
                            tm.InsertedOn = DateTime.Now;
                            tm.IsDeleted = 0;
                            tm.Mode = "BREAKDOWN";
                            tm.IsCompleted = 0;
                            db.TblSplivemodedb.Add(tm);
                            db.SaveChanges();

                            if (data.IndividualClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                                hmi.IsBreakdownClicked = 1;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            else if (data.OverallClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                                hmi.IsBreakdownClicked = 2;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            comobj.isTure = true;
                            comobj.response = "Red";
                        }
                        else { }

                    }
                    else if (data.Status == "End")
                    {
                        var hmibreakdown = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 1).FirstOrDefault();
                        if (hmibreakdown != null)
                        {
                            var tb = db.TblSpbreakdown.Where(m => m.BreakDownCode == data.LossCodeId && m.MachineId == data.MachineID && m.DoneWithRow == 0 && m.BatchNo == row).OrderByDescending(m => m.SpbreakdownId).FirstOrDefault();
                            if (tb != null)
                            {
                                tb.EndTime = DateTime.Now;
                                tb.DoneWithRow = 1;

                                db.Entry(tb).State = EntityState.Modified;
                                db.SaveChanges();
                            }


                            //get the latest row and update it.
                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.IsCompleted = 1;
                                modedata.EndTime = DateTime.Now;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.Entry(modedata).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            TblSplivemodedb tmIDLE = new TblSplivemodedb();
                            tmIDLE.ColorCode = "Green";
                            tmIDLE.BatchNumber = row;
                            tmIDLE.CorrectedDate = CorrectedDate;
                            var userid = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tmIDLE.InsertedBy = userid.UserId;
                            tmIDLE.InsertedOn = DateTime.Now;
                            tmIDLE.IsCompleted = 0;
                            tmIDLE.IsDeleted = 0;
                            tmIDLE.MachineId = data.MachineID;
                            tmIDLE.Mode = "PowerOn";
                            tmIDLE.StartTime = DateTime.Now;
                            db.TblSplivemodedb.Add(tmIDLE);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsBreakdownClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Individual BreakDown has not been started";
                                }
                            }
                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && (m.IsBreakdownClicked == 2 || m.IsBreakdownClicked == 1)).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsBreakdownClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Overall BreakDown has not been started";
                                }
                            }
                            comobj.isTure = true;
                            comobj.response = "Blue";
                        }
                        else
                        {
                            var tb = db.TblSpbreakdown.Where(m => m.BreakDownCode == data.LossCodeId && m.MachineId == data.MachineID && m.DoneWithRow == 0 && m.BatchNo == row).OrderByDescending(m => m.SpbreakdownId).FirstOrDefault();
                            if (tb != null)
                            {
                                tb.EndTime = DateTime.Now;
                                tb.DoneWithRow = 1;

                                db.Entry(tb).State = EntityState.Modified;
                                db.SaveChanges();
                            }


                            //get the latest row and update it.
                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.MachineID && m.IsCompleted == 0 && m.BatchNumber == row).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.IsCompleted = 1;
                                modedata.EndTime = DateTime.Now;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.Entry(modedata).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            TblSplivemodedb tmIDLE = new TblSplivemodedb();
                            tmIDLE.ColorCode = "Green";
                            tmIDLE.BatchNumber = row;
                            tmIDLE.CorrectedDate = CorrectedDate;
                            var userid = db.Tblusers.Where(m => m.MachineId == data.MachineID).FirstOrDefault();
                            tmIDLE.InsertedBy = userid.UserId;
                            tmIDLE.InsertedOn = DateTime.Now;
                            tmIDLE.IsCompleted = 0;
                            tmIDLE.IsDeleted = 0;
                            tmIDLE.MachineId = data.MachineID;
                            tmIDLE.Mode = "PowerOn";
                            tmIDLE.StartTime = DateTime.Now;
                            db.TblSplivemodedb.Add(tmIDLE);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsBreakdownClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Individual BreakDown has not been started";
                                }
                            }
                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row && m.IsBreakdownClicked == 2).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsBreakdownClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    comobj.isTure = false;
                                    comobj.response = "Overall BreakDown has not been started";
                                }
                            }
                            comobj.isTure = true;
                            comobj.response = "Blue";
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comobj.isTure = false;
            }
            return comobj;
        }

        public CommonResponse GenericWO(GenericPM data)
        {
            CommonResponse comobj = new CommonResponse();
            try
            {
                string[] hmiids = data.BatchNo.Split(',');
                foreach (var row in hmiids)
                {
                    if (data.Status == "Start") //Comes here while GenericWork is Started
                    {
                        //var operatorid = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == row).Select(m => m.OperatiorId).FirstOrDefault();
                        var breakdata = db.Tblgenericworkcodes.Where(m => m.IsDeleted == 0 && m.GenericWorkId == data.LossCodeId).FirstOrDefault();
                        string msgCode = breakdata.GenericWorkCode;
                        string msgDesc = breakdata.GenericWorkDesc;

                        TblSpGeneric tb = new TblSpGeneric();
                        tb.GenericCode = data.LossCodeId;
                        tb.OperatorId = data.opId;
                        tb.StartTime = DateTime.Now;
                        tb.BatchNumber = row;
                        tb.MachineId = data.MachineID;
                        tb.EndTime = null;
                        db.TblSpGeneric.Add(tb);
                        db.SaveChanges();

                        if (data.IndividualClicked == true)
                        {
                            var hmi1 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                            if (hmi1 != null)
                            {
                                hmi1.IsGenericClicked = 1;
                                db.Entry(hmi1).State = EntityState.Modified;
                                db.SaveChanges();
                            }
                            var SpHmidet = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == row).FirstOrDefault();
                            if (SpHmidet != null)
                            {

                                SpHmidet.IsGenericClicked = 1;
                                db.Entry(SpHmidet).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                        }
                        comobj.isTure = true;
                        comobj.response = "Green";
                    }
                    else if (data.Status == "End") // comes here while End GenericWork is Clicked
                    {
                        var tb = db.TblSpGeneric.Where(m => m.GenericCode == data.LossCodeId && m.MachineId == data.MachineID && m.BatchNumber == row).OrderByDescending(m => m.Id).FirstOrDefault();
                        if (tb != null)
                        {
                            tb.EndTime = DateTime.Now;
                            db.Entry(tb).State = EntityState.Modified;
                            db.SaveChanges();
                        }
                        if (data.IndividualClicked == false)
                        {
                            var hmi = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == row && m.IsGenericClicked == 1).FirstOrDefault();
                            if (hmi != null)
                            {
                                hmi.IsGenericClicked = 0;
                                db.Entry(hmi).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                        }
                        comobj.isTure = true;
                        comobj.response = "Blue";
                    }

                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                comobj.isTure = false;
            }
            return comobj;
        }

        public CommonResponse GetActivity(int processId)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<Activity> actobj = new List<Activity>();
                var act = db.Tblactivity.Where(m => m.Isdeleted == 0 && m.ProcessId == processId).ToList();
                foreach (var row in act)
                {
                    Activity obj1 = new Activity();
                    obj1.ActivityId = row.ActivityId;
                    obj1.ActivityName = row.ActivityName;
                    actobj.Add(obj1);
                }
                if (actobj.Count != 0)
                {
                    obj.isTure = true;
                    obj.response = actobj;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponse GetProcess()
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<process> actobj = new List<process>();
                var act = db.TblProcess.Where(m => m.Isdeleted == 0).ToList();
                foreach (var row in act)
                {
                    process obj1 = new process();
                    obj1.processid = row.ProcessId;
                    obj1.processdesc = row.ProcessName;
                    actobj.Add(obj1);
                }
                if (actobj.Count != 0)
                {
                    obj.isTure = true;
                    obj.response = actobj;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        //set Delivered Qty
        //public CommonResponsewithEror SetDeliveredQty(SetDel data)
        //{
        //    CommonResponsewithEror obj = new CommonResponsewithEror();
        //    try
        //    {
        //        var HmiData = db.TblSplivehmiscreen.Where(x => x.Sphmiid == data.HMIID).FirstOrDefault();
        //        if (HmiData != null)
        //        {
        //            HmiData.DeliveredQty = data.DeliveredQty;
        //            db.SaveChanges();

        //            HmiData = db.TblSplivehmiscreen.Where(x => x.Sphmiid == data.HMIID).FirstOrDefault();

        //            int deliveredQty = 0;
        //            if (int.TryParse(Convert.ToString(HmiData.DeliveredQty), out deliveredQty))
        //            {
        //                int processed = 0;
        //                int.TryParse(Convert.ToString(HmiData.ProcessQty), out processed);
        //                if ((deliveredQty + processed) > Convert.ToInt32(HmiData.TargetQty))
        //                {
        //                    obj.isTure = false;
        //                    obj.errorMsg = "DeliveredQty Must be less than Target Quantity for " + HmiData.WorkOrderNo;
        //                    HmiData.DeliveredQty = 0;
        //                    db.SaveChanges();
        //                }
        //                //else
        //                //{
        //                //    HmiData.DeliveredQty =0;
        //                //    db.SaveChanges();
        //                //}
        //            }
        //            obj.isTure = true;
        //            obj.response = ResourceResponse.UpdatedSuccessMessage;
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

        public CommonResponsewithEror SetDeliveredQty(SetDel data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            try
            {
                var HmiData = db.TblSplivehmiscreen.Where(x => x.Sphmiid == data.HMIID).FirstOrDefault();
                if (HmiData != null)
                {
                    HmiData.DeliveredQty = data.DeliveredQty;
                    HmiData.RejectionQuantity = data.RejectionQty;
                    db.SaveChanges();

                    HmiData = db.TblSplivehmiscreen.Where(x => x.Sphmiid == data.HMIID).FirstOrDefault();

                    int deliveredQty = 0;
                    if (int.TryParse(Convert.ToString(HmiData.DeliveredQty), out deliveredQty))
                    {
                        int processed = 0;
                        int.TryParse(Convert.ToString(HmiData.ProcessQty), out processed);

                        int rejInt = 0;
                        int.TryParse(Convert.ToString(HmiData.RejectionQuantity), out rejInt);


                        if ((deliveredQty + processed + rejInt) > Convert.ToInt32(HmiData.TargetQty))
                        {
                            obj.isTure = false;
                            obj.errorMsg = "DeliveredQty and RejectionQty Must be less than Target Quantity for " + HmiData.WorkOrderNo;
                            HmiData.DeliveredQty = 0;
                            HmiData.RejectionQuantity = 0;
                            db.SaveChanges();
                        }
                        //else
                        //{
                        //    HmiData.DeliveredQty =0;
                        //    db.SaveChanges();
                        //}

                        else
                        {
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;

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
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponse SetShift(SetShiftbatch data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] batchno = data.BatchNo.Split(',');
                foreach (var batchrow in batchno)
                {
                    string Batchno = Convert.ToString(batchrow);
                    var HmiData = db.TblSplivehmiscreen.Where(x => x.AutoBatchNumber == batchrow).ToList();
                    if (HmiData != null)
                    {
                        foreach (var row in HmiData)
                        {
                            row.Shift = data.Shift;
                            row.PestartTime = DateTime.Now;
                            row.IsUpdate = 1;
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }

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
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponse SetChange(string batchNo)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string[] HMIIDArray = batchNo.Split(',');

                //2)
                int ExceptionHMIID = 0;
                foreach (var hmiid in HMIIDArray)
                {
                    var HmiData = db.TblSplivehmiscreen.Where(x => x.AutoBatchNumber == hmiid).ToList();
                    if (HmiData != null)
                    {
                        foreach (var row in HmiData)
                        {
                            row.IsUpdate = 0;
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }

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
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        public CommonResponsewithEror split(splitWo data)
        {
            CommonResponsewithEror comres = new CommonResponsewithEror();
            try
            {
                var HMIData = db.TblSplivehmiscreen.Where(m => m.Sphmiid == data.HMIID).FirstOrDefault();
                if (HMIData != null)
                {
                    if (HMIData.Date != null)
                    {
                        int deliveredQty = 0;
                        if (int.TryParse(Convert.ToString(HMIData.DeliveredQty), out deliveredQty))
                        {
                            int processed = 0;
                            int.TryParse(Convert.ToString(HMIData.ProcessQty), out processed);
                            if ((deliveredQty + processed) >= Convert.ToInt32(HMIData.TargetQty))
                            {
                                comres.isTure = false;
                                comres.errorMsg = "DeliveredQty Must be less than Target Quantity for " + HMIData.WorkOrderNo;
                            }
                            else if (data.isChecked == true)
                            {
                                HMIData.SplitWo = "Yes";
                                db.SaveChanges();
                                comres.isTure = true;
                                comres.response = ResourceResponse.NoItemsFound;
                            }
                            else if (data.isChecked == false)
                            {
                                HMIData.SplitWo = "No";
                                db.SaveChanges();
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        //Anjali Code

        #region     

        public CommonResponse BatchUniqueCodeGenerator(string workCenterName)
        {
            CommonResponse obj = new CommonResponse();
            try
            {

                var uniqueCode = " ";
                DateTime date = DateTime.Now;
                string[] time = date.ToString().Split(' ');
                string newPathTime = time[1].Replace(":", string.Empty).ToString();
                newPathTime = newPathTime.Remove(newPathTime.Length - 2, 2);
                string newDate = date.ToShortDateString().ToString();
                newDate = newDate.Replace("-", string.Empty);
                newDate = newDate.Replace("/", string.Empty);
                uniqueCode = workCenterName + newDate + newPathTime;
                obj.isTure = true;
                obj.response = uniqueCode;

            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponse NewBatchUniqueCodeGenerator(int MachineId)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var machineIncNo = db.Tblmachinedetails.Where(m => m.MachineId == MachineId && m.IsDeleted == 0).Select(m => m.MachineInvNo).FirstOrDefault();
                var uniqueCode = " ";
                DateTime date = DateTime.Now;
                string[] time = date.ToString().Split(' ');
                string newPathTime = time[1].Replace(":", string.Empty).ToString();
                newPathTime = newPathTime.Remove(newPathTime.Length - 2, 2);
                string newDate = date.ToShortDateString().ToString();
                newDate = newDate.Replace("-", string.Empty);
                newDate = newDate.Replace("/", string.Empty);
                uniqueCode = machineIncNo + newDate + newPathTime;
                obj.isTure = true;
                obj.response = uniqueCode;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        //public CommonResponse AddBatchDetails(BatchDetails data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        string getDate = DateTime.Now.ToString("yyyy-MM-dd");
        //        string[] aary = data.DdlIds.Split(',');
        //        List<int> intArry = aary.ToList().Select(int.Parse).ToList();
        //        var DDLData = db.Tblddl.Where(m => intArry.Contains(m.Ddlid)).ToList();
        //        bool isValid = true;
        //        foreach (var item in DDLData)
        //        {
        //            int DDLID = item.Ddlid;
        //            var ddldataInner = db.Tblddl.Where(m => m.IsCompleted == 0 && m.Ddlid == DDLID && m.IsWoselected == 0).FirstOrDefault();
        //            string SplitWOInner = ddldataInner.SplitWo;
        //            string WONoInner = ddldataInner.WorkOrder;
        //            string PartInner = ddldataInner.MaterialDesc;
        //            string OperationInner = ddldataInner.OperationNo;
        //            string Type = ddldataInner.Type;

        //            var HMICheck = db.Tblhmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner).FirstOrDefault();
        //            if(HMICheck != null)
        //            {
        //                if(HMICheck.Date != null && HMICheck.Time == null)
        //                {
        //                    obj.isTure = false;
        //                    obj.response = "Work Order Already Started";
        //                }
        //            }

        //            var HMIBatchCheck = db.Tblbatchhmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner).FirstOrDefault();
        //            if(HMIBatchCheck != null)
        //            {
        //                if(HMIBatchCheck.Date == null)
        //                {
        //                    obj.isTure = false;
        //                    obj.response = "Work Order Already Selected";
        //                }
        //            }

        //            var check = db.Tblbatchhmiscreen.Where(m => m.OperationNo == item.OperationNo && m.WorkOrderNo == item.Work_Order_No && m.PartNo == item.PartNo && m.IsWorkInProgress != 2).FirstOrDefault();
        //            if (check == null)
        //            {
        //                Tblbatchhmiscreen tblbatchhmiscreen = new Tblbatchhmiscreen();
        //                tblbatchhmiscreen.OperatiorId = 0;
        //                tblbatchhmiscreen.Status = 0;
        //                tblbatchhmiscreen.Project = item.Project;
        //                tblbatchhmiscreen.PartNo = item.PartNo;
        //                tblbatchhmiscreen.OperationNo = item.OperationNo;
        //                tblbatchhmiscreen.WorkOrderNo = item.Work_Order_No;
        //                tblbatchhmiscreen.TargetQty = item.TargetQty;
        //                tblbatchhmiscreen.CorrectedDate = getDate;
        //                tblbatchhmiscreen.ProdFai = item.Prod_FAI;
        //                tblbatchhmiscreen.PestartTime = DateTime.Now;
        //                tblbatchhmiscreen.DdlwokrCentre = item.DDLWokrCentre;
        //                tblbatchhmiscreen.AutoBatchNumber = item.UniqueBatchNo;
        //                tblbatchhmiscreen.MachineId = item.MachineId;
        //                db.Tblbatchhmiscreen.Add(tblbatchhmiscreen);
        //                db.SaveChanges();
        //                obj.isTure = true;
        //            }
        //            else
        //            {
        //                obj.isTure = false;
        //                obj.response = "Work Order Already Exist";
        //            }
        //        }
        //        if (obj.isTure == true)
        //        {
        //            var dbCheck = (from wf in db.Tblbatchhmiscreen
        //                           where wf.Date == null
        //                           select new
        //                           {
        //                               autoBatchNumber = wf.AutoBatchNumber
        //                           }).ToList().Distinct();
        //            if (dbCheck.Count() != 0)
        //            {
        //                obj.isTure = true;
        //                obj.response = dbCheck;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isTure = false;
        //    }
        //    return obj;
        //}

        //public CommonResponse GetHoldCodes(int HoldCodeID)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        if (HoldCodeID != 0)
        //        {
        //            var check = db.Tblholdcodes.Where(m => m.HoldCodeId == HoldCodeID).FirstOrDefault();
        //            var level = check.HoldCodesLevel;
        //            string lossCodes = check.HoldCode;
        //            if (level == 1)
        //            {
        //                var level2Data = db.Tblholdcodes.Where(m => m.IsDeleted == 0 && m.HoldCodesLevel1Id == HoldCodeID && m.HoldCodesLevel == 2 && m.HoldCodesLevel2Id == null).ToList();

        //                if (level2Data.Count == 0)
        //                {
        //                    var level1Data = db.Tblholdcodes.Where(m => m.IsDeleted == 0 && m.HoldCodesLevel == level && m.HoldCodesLevel1Id == null && m.HoldCodesLevel2Id == null).ToList();
        //                    obj.isTure = false;
        //                    obj.response = level1Data;
        //                }
        //                obj.isTure = true;
        //                obj.response = level2Data;
        //            }
        //            else if (level == 2)
        //            {
        //                var level3Data = db.Tblholdcodes.Where(m => m.IsDeleted == 0 && m.HoldCodesLevel2Id == HoldCodeID && m.HoldCodesLevel == 3).ToList();
        //                int prevLevelId = Convert.ToInt32(check.HoldCodesLevel1Id);
        //                var level1data = db.Tblholdcodes.Where(m => m.HoldCodeId == prevLevelId).Select(m => m.HoldCode).FirstOrDefault();
        //                if (level3Data.Count == 0)
        //                {
        //                    var level2Data = db.Tblholdcodes.Where(m => m.IsDeleted == 0 && m.HoldCodesLevel1Id == prevLevelId && m.HoldCodesLevel2Id == null).ToList();
        //                    obj.isTure = false;
        //                    obj.response = level2Data;
        //                }
        //                obj.isTure = true;
        //                obj.response = level3Data;
        //            }
        //            else if (level == 3)
        //            {
        //                int prevLevelId = Convert.ToInt32(check.HoldCodesLevel2Id);
        //                int FirstLevelID = Convert.ToInt32(check.HoldCodesLevel1Id);
        //                var level2scrum = db.Tblholdcodes.Where(m => m.HoldCodeId == prevLevelId).Select(m => m.HoldCode).FirstOrDefault();
        //                var level1scrum = db.Tblholdcodes.Where(m => m.HoldCodeId == FirstLevelID).Select(m => m.HoldCode).FirstOrDefault();
        //                var level2Data = db.Tblholdcodes.Where(m => m.IsDeleted == 0 && m.HoldCodesLevel2Id == prevLevelId && m.HoldCodesLevel == 3).ToList();
        //                obj.isTure = true;
        //                obj.response = level2Data;
        //            }
        //        }
        //        else
        //        {
        //            var check = (from wf in db.Tblholdcodes
        //                         where wf.HoldCodesLevel == 1 && wf.IsDeleted == 0
        //                         select new
        //                         {
        //                             HoldCode = wf.HoldCode,
        //                             HoldCodeDesc = wf.HoldCodeDesc,
        //                             HoldCodeId = wf.HoldCodeId
        //                         }).ToList();
        //            if (check.Count != 0)
        //            {
        //                obj.isTure = true;
        //                obj.response = check;
        //            }
        //            else
        //            {
        //                obj.isTure = false;
        //                obj.response = "No Items Found";
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isTure = false;
        //    }
        //    return obj;
        //}

        public string[] GetDateShift()
        {
            string[] dateShift = new string[2];
            //Get CorrectedDate & shift
            #region
            string Shift = null;
            var queryshift = (from wf in db.TblshiftMstr
                              where wf.IsDeleted == 0
                              select new
                              {
                                  ShiftName = wf.ShiftName,
                                  StartTime = wf.StartTime,
                                  EndTime = wf.EndTime
                              }).ToList();

            if (queryshift != null)
            {
                String[] msgtime = System.DateTime.Now.TimeOfDay.ToString().Split(':');
                TimeSpan msgstime = System.DateTime.Now.TimeOfDay;
                //TimeSpan msgstime = new TimeSpan(Convert.ToInt32(msgtime[0]), Convert.ToInt32(msgtime[1]), Convert.ToInt32(msgtime[2]));
                TimeSpan s1t1 = new TimeSpan(0, 0, 0), s1t2 = new TimeSpan(0, 0, 0), s2t1 = new TimeSpan(0, 0, 0), s2t2 = new TimeSpan(0, 0, 0);
                TimeSpan s3t1 = new TimeSpan(0, 0, 0), s3t2 = new TimeSpan(0, 0, 0), s3t3 = new TimeSpan(0, 0, 0), s3t4 = new TimeSpan(23, 59, 59);
                foreach (var item in queryshift)
                {
                    if (item.ShiftName.ToString().Contains("A"))
                    {
                        String[] s1 = item.StartTime.ToString().Split(':');
                        s1t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                        String[] s11 = item.EndTime.ToString().Split(':');
                        s1t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                    }
                    else if (item.ShiftName.ToString().Contains("B"))
                    {
                        String[] s1 = item.StartTime.ToString().Split(':');
                        s2t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                        String[] s11 = item.EndTime.ToString().Split(':');
                        s2t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                    }
                    else if (item.ShiftName.ToString().Contains("C"))
                    {
                        String[] s1 = item.StartTime.ToString().Split(':');
                        s3t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                        String[] s11 = item.EndTime.ToString().Split(':');
                        s3t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                    }
                }
                String CorrectedDate = System.DateTime.Now.ToString("yyyy-MM-dd");
                if (msgstime >= s1t1 && msgstime < s1t2)
                {
                    Shift = "A";
                }
                else if (msgstime >= s2t1 && msgstime < s2t2)
                {
                    Shift = "B";
                }
                else if ((msgstime >= s3t1 && msgstime <= s3t4) || (msgstime >= s3t3 && msgstime < s3t2))
                {
                    Shift = "C";
                    if (msgstime >= s3t3 && msgstime < s3t2)
                    {
                        CorrectedDate = System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                    }
                }
                #endregion
                dateShift[0] = Shift;
                dateShift[1] = CorrectedDate;
            }
            return dateShift;
        }

        /// <summary>
        /// Hold Code Entry
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public GeneralResponse HoldCodeEntry(AddHoldCodes data)
        {
            GeneralResponse obj = new GeneralResponse();
            try
            {
                string[] ids = data.Hmiid.Split(',');
                int[] myInts = Array.ConvertAll(ids, int.Parse);
                #region
                string CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                if (DateTime.Now.Hour < 6 && DateTime.Now.Hour >= 0)
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                #endregion

                foreach (var item in myInts)
                {
                    var tblhmi = db.TblSplivehmiscreen.Where(m => m.Sphmiid == item).FirstOrDefault();
                    if (tblhmi != null)
                    {
                        tblhmi.Time = DateTime.Now;
                        tblhmi.IsWorkInProgress = 0;
                        tblhmi.IsHold = 1;
                        tblhmi.Status = 2;
                        db.SaveChanges();
                    }

                    var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == tblhmi.Sphmiid).Select(m => m.AutoBatchNumber).FirstOrDefault();
                    var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                    hmi.IsHoldClicked = 1;
                    db.Entry(hmi).State = EntityState.Modified;
                    db.SaveChanges();

                    var holdCodeData = db.Tblholdcodes.Where(m => m.HoldCodeId == data.HoldCodeId).FirstOrDefault();

                    TblSplivemanuallossofentry tblSplivemanuallossofentry = new TblSplivemanuallossofentry();
                    tblSplivemanuallossofentry.CorrectedDate = CorrectedDate;
                    tblSplivemanuallossofentry.Hmiid = item;
                    tblSplivemanuallossofentry.MessageCodeId = data.HoldCodeId;
                    tblSplivemanuallossofentry.MessageDesc = holdCodeData.HoldCodeDesc;
                    tblSplivemanuallossofentry.MessageCode = holdCodeData.HoldCode;
                    tblSplivemanuallossofentry.Wono = tblhmi.WorkOrderNo;
                    tblSplivemanuallossofentry.OpNo = Convert.ToInt32(tblhmi.OperationNo);
                    tblSplivemanuallossofentry.PartNo = tblhmi.PartNo;
                    tblSplivemanuallossofentry.StartDateTime = DateTime.Now;
                    string[] GetDateShift1 = GetDateShift();
                    tblSplivemanuallossofentry.Shift = GetDateShift1[0];
                    tblSplivemanuallossofentry.MachineId = data.MachineId;
                    db.TblSplivemanuallossofentry.Add(tblSplivemanuallossofentry);
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Added Successfully";
                    obj.color = "Yellow";
                }
                obj.isStatus = true;
                obj.response = "Yellow";
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        public GeneralResponse EndHold(string hmiiid)
        {
            GeneralResponse obj = new GeneralResponse();
            try
            {
                string[] ids = hmiiid.Split(',');
                foreach (var item in ids)
                {
                    DateTime EndTime = DateTime.Now;
                    int HMIID = Convert.ToInt32(item);
                    var hmiData = db.Tbllivehmiscreen.Where(m => m.Hmiid == HMIID).FirstOrDefault();
                    if (hmiData != null)
                    {
                        hmiData.IsHold = 2;
                        db.SaveChanges();
                    }

                    var batchno = db.TblSplivehmiscreen.Where(m => m.Sphmiid == HMIID).Select(m => m.AutoBatchNumber).FirstOrDefault();
                    var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == batchno).FirstOrDefault();
                    hmi.IsHoldClicked = 0;
                    db.Entry(hmi).State = EntityState.Modified;
                    db.SaveChanges();

                    var tblmanualLossData = db.Tbllivemanuallossofentry.Where(m => m.Hmiid == HMIID).OrderByDescending(m => m.StartDateTime).FirstOrDefault();
                    if (tblmanualLossData != null)
                    {
                        tblmanualLossData.EndHmiid = HMIID;
                        tblmanualLossData.EndDateTime = EndTime;
                        db.SaveChanges();
                        obj.isStatus = true;
                        obj.response = "Updated Successfully";
                    }
                }
                obj.isStatus = true;
                obj.response = "Blue";
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        //public CommonResponse GetIdleCodes(int LossCodeID, bool isStart)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        if (LossCodeID != 0)
        //        {
        //            var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == LossCodeID).FirstOrDefault();
        //            int level = lossdata.LossCodesLevel;
        //            string losscode = lossdata.LossCode;
        //            if (level == 1)
        //            {
        //                var level2Data = (from wf in db.Tbllossescodes
        //                                  where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == LossCodeID && wf.LossCodesLevel == 2 && wf.LossCodesLevel2Id == null && wf.MessageType != "BREAKDOWN"
        //                                  select new
        //                                  {
        //                                      LossCodeId = wf.LossCodeId,
        //                                      LossCode = wf.LossCode,
        //                                      LossCodeDesc = wf.LossCodeDesc,
        //                                      MessageType = wf.MessageType,
        //                                      isStart = isStart
        //                                  }).ToList();
        //                if (level2Data.Count == 0)
        //                {
        //                    var level1Data = (from wf in db.Tbllossescodes
        //                                      where wf.IsDeleted == 0 && wf.LossCodesLevel == level && wf.LossCodesLevel1Id == null && wf.LossCodesLevel2Id == null && wf.MessageType != "NoCode" && wf.MessageType != "BREAKDOWN" && wf.MessageType != "PM"
        //                                      select new
        //                                      {
        //                                          LossCodeId = wf.LossCodeId,
        //                                          LossCode = wf.LossCode,
        //                                          LossCodeDesc = wf.LossCodeDesc,
        //                                          MessageType = wf.MessageType,
        //                                          isStart = isStart
        //                                      }).ToList();
        //                    obj.isTure = true;
        //                    obj.response = level1Data;
        //                }
        //                else
        //                {
        //                    obj.isTure = true;
        //                    obj.response = level2Data;
        //                }
        //            }
        //            else if (level == 2)
        //            {
        //                var level3Data = (from wf in db.Tbllossescodes
        //                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == LossCodeID && wf.LossCodesLevel == 3 && wf.MessageType != "BREAKDOWN"
        //                                  select new
        //                                  {
        //                                      LossCodeId = wf.LossCodeId,
        //                                      LossCode = wf.LossCode,
        //                                      LossCodeDesc = wf.LossCodeDesc,
        //                                      MessageType = wf.MessageType,
        //                                      isStart = isStart
        //                                  }).ToList();
        //                int prevLevelId = Convert.ToInt32(lossdata.LossCodesLevel1Id);
        //                if (level3Data.Count == 0)
        //                {
        //                    var level2Data = (from wf in db.Tbllossescodes
        //                                      where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == prevLevelId && wf.LossCodesLevel2Id == null
        //                                      select new
        //                                      {
        //                                          LossCodeId = wf.LossCodeId,
        //                                          LossCode = wf.LossCode,
        //                                          LossCodeDesc = wf.LossCodeDesc,
        //                                          MessageType = wf.MessageType,
        //                                          isStart = isStart
        //                                      }).ToList();
        //                    obj.isTure = true;
        //                    obj.response = level2Data;
        //                }
        //                else
        //                {
        //                    obj.isTure = true;
        //                    obj.response = level3Data;
        //                }
        //            }
        //            else if (level == 3)
        //            {
        //                int prevLevelId = Convert.ToInt32(lossdata.LossCodesLevel2Id);
        //                var level2Data = (from wf in db.Tbllossescodes
        //                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == prevLevelId && wf.LossCodesLevel == 3
        //                                  select new
        //                                  {
        //                                      LossCodeId = wf.LossCodeId,
        //                                      LossCode = wf.LossCode,
        //                                      LossCodeDesc = wf.LossCodeDesc,
        //                                      MessageType = wf.MessageType,
        //                                      isStart = isStart
        //                                  }).ToList();
        //                obj.isTure = true;
        //                obj.response = level2Data;
        //            }
        //        }
        //        else
        //        {
        //            var level1Data = (from wf in db.Tbllossescodes
        //                              where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType != "NoCode" && wf.MessageType != "BREAKDOWN" && wf.MessageType != "PM"
        //                              select new
        //                              {
        //                                  LossCodeId = wf.LossCodeId,
        //                                  LossCode = wf.LossCode,
        //                                  LossCodeDesc = wf.LossCodeDesc,
        //                                  MessageType = wf.MessageType,
        //                                  isStart = isStart
        //                              }).ToList();

        //            if (level1Data.Count != 0)
        //            {
        //                obj.isTure = true;
        //                obj.response = level1Data;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isTure = false;
        //    }
        //    return obj;
        //}

        public CommonResponse IdleLossCodes(AddIdleCodes data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string CorrectedDate = null;
                Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan Start = StartTime.StartTime;
                if (Start <= DateTime.Now.TimeOfDay)
                {
                    CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }

                DateTime Time = DateTime.Now;
                TimeSpan Tm = new TimeSpan(Time.Hour, Time.Minute, Time.Second);
                var ShiftDetails = db.TblshiftMstr.Where(m => m.StartTime <= Tm && m.EndTime >= Tm).FirstOrDefault();
                string Shift = "C";
                if (ShiftDetails != null)
                {
                    Shift = ShiftDetails.ShiftName;
                }
                if (data.LossCodeID != 0 && data.isStart == true)
                {
                    var breakdata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeId == data.LossCodeID).FirstOrDefault();
                    string msgCode = breakdata.LossCode;
                    string msgDesc = breakdata.LossCodeDesc;

                    Tbllivelossofentry tbllivelossofentry = new Tbllivelossofentry();
                    if (data.LossCodeID != 0)
                    {
                        var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.LossCodeID).FirstOrDefault();
                        tbllivelossofentry.MessageCodeId = data.LossCodeID;
                        tbllivelossofentry.MessageDesc = lossdata.LossCodeDesc;
                        tbllivelossofentry.MessageCode = lossdata.LossCode;
                    }
                    else
                    {
                        int MessageCodeID = Convert.ToInt32(tbllivelossofentry.MessageCodeId);
                        tbllivelossofentry.MessageCodeId = Convert.ToInt32(tbllivelossofentry.MessageCodeId);
                        var a = db.MessageCodeMaster.Where(m => m.MessageCodeId == MessageCodeID).FirstOrDefault();
                        tbllivelossofentry.MessageDesc = a.MessageDescription.ToString();
                        tbllivelossofentry.MessageCode = a.MessageCode.ToString();
                    }

                    tbllivelossofentry.StartDateTime = DateTime.Now;
                    tbllivelossofentry.EntryTime = DateTime.Now;
                    tbllivelossofentry.CorrectedDate = CorrectedDate;
                    tbllivelossofentry.MachineId = data.MachineId;
                    tbllivelossofentry.Shift = Shift;
                    tbllivelossofentry.MessageCodeId = tbllivelossofentry.MessageCodeId;
                    tbllivelossofentry.MessageDesc = tbllivelossofentry.MessageDesc;
                    tbllivelossofentry.MessageCode = tbllivelossofentry.MessageCode;
                    tbllivelossofentry.IsUpdate = 1;
                    tbllivelossofentry.DoneWithRow = 0;
                    tbllivelossofentry.IsStart = 1;
                    tbllivelossofentry.IsScreen = 0;
                    tbllivelossofentry.ForRefresh = 1;
                    db.Tbllivelossofentry.Add(tbllivelossofentry);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Added Successfully";

                    var modedata = db.Tbllivemodedb.Where(m => m.MachineId == data.MachineId && m.IsCompleted == 0).OrderByDescending(m => m.StartTime).FirstOrDefault();
                    if (modedata != null)
                    {
                        modedata.IsCompleted = 1;
                        modedata.EndTime = DateTime.Now;
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Updated Successfully";
                    }

                    Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                    tbllivemodedb.MachineId = data.MachineId;
                    tbllivemodedb.CorrectedDate = CorrectedDate;
                    tbllivemodedb.StartTime = DateTime.Now;
                    tbllivemodedb.ColorCode = "yellow";
                    tbllivemodedb.InsertedOn = DateTime.Now;
                    tbllivemodedb.IsDeleted = 0;
                    tbllivemodedb.Mode = "IDLE";
                    tbllivemodedb.IsCompleted = 0;
                    db.Tbllivemodedb.Add(tbllivemodedb);
                    db.SaveChanges();

                    obj.isTure = true;
                    obj.response = "Yellow";
                }
                else if (data.LossCodeID != 0 && data.isStart == false)
                {
                    var tb = db.Tbllivelossofentry.Where(m => m.MessageCodeId == data.LossCodeID && m.MachineId == data.MachineId && m.DoneWithRow == 0).OrderByDescending(m => m.LossId).FirstOrDefault();
                    if (tb != null)
                    {
                        tb.EndDateTime = DateTime.Now;
                        tb.DoneWithRow = 1;
                        tb.IsUpdate = 1;
                        tb.IsScreen = 0;
                        tb.IsStart = 0;
                        tb.ForRefresh = 0;
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Updated Successfully";
                    }
                    var modedata = db.Tbllivemodedb.Where(m => m.MachineId == data.MachineId && m.IsCompleted == 0).OrderByDescending(m => m.StartTime).FirstOrDefault();
                    if (modedata != null)
                    {
                        modedata.EndTime = DateTime.Now;
                        modedata.IsCompleted = 1;
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Updated Successfully";
                    }

                    Tbllivemodedb tbllivemodedb = new Tbllivemodedb();
                    tbllivemodedb.MachineId = data.MachineId;
                    tbllivemodedb.CorrectedDate = CorrectedDate;
                    tbllivemodedb.StartTime = DateTime.Now;
                    tbllivemodedb.ColorCode = "green";
                    tbllivemodedb.InsertedOn = DateTime.Now;
                    tbllivemodedb.IsDeleted = 0;
                    tbllivemodedb.Mode = "PowerOn";
                    tbllivemodedb.IsCompleted = 0;
                    db.Tbllivemodedb.Add(tbllivemodedb);
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = "Blue";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// Idle List
        /// </summary>
        /// <param name="MachineId"></param>
        /// <returns></returns>
        /// <summary>
        /// Idle List
        /// </summary>
        /// <param name="MachineId"></param>
        /// <returns></returns>
        public CommonResponse IdleList(int MachineId = 0)
        {
            CommonResponse obj = new CommonResponse();
            try
            {

                string CorrectedDate = null;

                Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan Start = StartTime.StartTime;
                if (Start <= DateTime.Now.TimeOfDay)
                {
                    CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                string[] DateShiftValues = GetDateShift();
                CorrectedDate = DateShiftValues[1];
                var machinedispname = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == MachineId).Select(m => m.MachineDispName).FirstOrDefault();

                var check = (from wf in db.TblSplivelossofentry
                             join lc in db.Tbllossescodes on wf.MessageCodeId equals lc.LossCodeId
                             where wf.DoneWithRow == 1 && wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate
                             select new
                             {
                                 LossCodeId = lc.LossCodeId,
                                 LossCode = lc.LossCode,
                                 LossCodesLevel = lc.LossCodesLevel,
                                 LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                 LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                 StartDateTime = Convert.ToDateTime(wf.StartDateTime).ToString("yyyy-MM-dd HH:mm:ss"),
                                 EndDateTime = Convert.ToDateTime(wf.EndDateTime).ToString("yyyy-MM-dd HH:mm:ss"),
                                 MachineId = wf.MachineId,
                                 MessageCodeId = wf.MessageCodeId,
                                 CorrectedDate = wf.CorrectedDate,
                                 SplossId = wf.SplossId,


                                 ress1 = (lc.LossCodesLevel == 1 ?
                                         lc.LossCode :
                                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                 ress2 = (lc.LossCodesLevel == 2 ?
                                         lc.LossCode :
                                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                 ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),

                                 duration = Convert.ToDateTime(wf.EndDateTime).Subtract(Convert.ToDateTime(wf.StartDateTime)).TotalSeconds,

                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isTure = true;
                    obj.response = check;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No Items Found";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// Break Down List
        /// </summary>
        /// <param name="MachineId"></param>
        /// <returns></returns>
        public CommonResponse BreakDownList(int MachineId = 0)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                string CorrectedDate = null;
                Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan Start = StartTime.StartTime;
                if (Start <= DateTime.Now.TimeOfDay)
                {
                    CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                var machinedispname = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == MachineId).Select(m => m.MachineDispName).FirstOrDefault();

                var check = (from wf in db.TblSpbreakdown
                             join lc in db.Tbllossescodes on wf.BreakDownCode equals lc.LossCodeId
                             where wf.DoneWithRow == 1 && wf.MachineId == MachineId && wf.CorrectedDate == CorrectedDate
                             select new
                             {
                                 LossCodeId = lc.LossCodeId,
                                 LossCode = lc.LossCode,
                                 LossCodesLevel = lc.LossCodesLevel,
                                 LossCodesLevel1Id = lc.LossCodesLevel1Id,
                                 LossCodesLevel2Id = lc.LossCodesLevel2Id,
                                 StartDateTime = Convert.ToDateTime(wf.StartTime).ToString("yyyy-MM-dd HH:mm:ss"),
                                 EndDateTime = Convert.ToDateTime(wf.EndTime).ToString("yyyy-MM-dd HH:mm:ss"),
                                 MachineId = wf.MachineId,
                                 MessageCodeId = wf.BreakDownCode,
                                 CorrectedDate = wf.CorrectedDate,
                                 SplossId = wf.SpbreakdownId,


                                 ress1 = (lc.LossCodesLevel == 1 ?
                                         lc.LossCode :
                                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel1Id).Select(m => m.LossCode).FirstOrDefault()),


                                 ress2 = (lc.LossCodesLevel == 2 ?
                                         lc.LossCode :
                                         db.Tbllossescodes.Where(m => m.LossCodeId == lc.LossCodesLevel2Id).Select(m => m.LossCode).FirstOrDefault()),

                                 ress3 = (lc.LossCodesLevel == 3 ? lc.LossCode : null),


                                 duration = Convert.ToDateTime(wf.EndTime).Subtract(Convert.ToDateTime(wf.StartTime)).TotalSeconds,

                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isTure = true;
                    obj.response = check;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public AddWOCommonResponse AddUploadedDdlBatchDetails(BatchDetails data)
        {
            AddWOCommonResponse obj = new AddWOCommonResponse();
            try
            {
                string getDate = DateTime.Now.ToString("yyyy-MM-dd");
                string[] aary = data.DdlIds.Split(',');
                List<int> intArry = aary.ToList().Select(int.Parse).ToList();
                var DDLData = db.TblSpddl.Where(m => intArry.Contains(m.SpDdlid)).ToList();
                List<Responce> responces = new List<Responce>();

                string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                SqlConnection conn = new SqlConnection(connectionstring);

                SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblspddl]", conn);
                conn.Open();
                truncateQuery.ExecuteNonQuery();
                conn.Close();

                foreach (var item in DDLData)
                {
                    Responce responce = new Responce();
                    int DDLID = item.SpDdlid;
                    var ddldataInner = db.TblSpddl.Where(m => m.IsCompleted == 0 && m.SpDdlid == DDLID && m.IsWoselected == 0).FirstOrDefault();
                    bool isValid = true;
                    if (ddldataInner != null)
                    {
                        string SplitWOInner = ddldataInner.SplitWo;
                        string WONoInner = ddldataInner.WorkOrder;
                        string PartInner = ddldataInner.MaterialDesc;
                        string OperationInner = ddldataInner.OperationNo;
                        string Type = ddldataInner.Type;

                        var HMICheck = db.TblSplivehmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner).FirstOrDefault();
                        if (HMICheck != null)
                        {
                            if (HMICheck.Date != null && HMICheck.Time == null)
                            {
                                responce.message = "WorkOrderNo:" + WONoInner + "and PartNo:" + PartInner + "and OperationNo:" + OperationInner + " Work Order Already Started";
                                responces.Add(responce);
                                isValid = false;
                            }
                        }

                        var HMIBatchCheck = db.Tblbatchhmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner).FirstOrDefault();
                        if (HMIBatchCheck != null)
                        {
                            if (HMIBatchCheck.Date == null)
                            {
                                obj.isStatus = false;
                                responce.message = "WorkOrderNo:" + WONoInner + "and PartNo:" + PartInner + "and OperationNo:" + OperationInner + "  Work Order Already Selected";
                                responces.Add(responce);
                                isValid = false;
                            }
                        }

                        var HMISequenceCheck = db.TblSplivehmiscreen.Where(m => m.WorkOrderNo == WONoInner && m.OperationNo == OperationInner && m.PartNo == PartInner).Select(m => m.OperationNo).FirstOrDefault();
                        int OperationNo = Convert.ToInt32(HMISequenceCheck);
                        if (HMISequenceCheck != null)
                        {
                            var dbCheck = db.Tblddl.Where(m => m.WorkOrder == WONoInner && Convert.ToInt32(m.OperationNo) > OperationNo).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                obj.isStatus = false;
                                responce.message = "Select WorkOrderNo:" + dbCheck.WorkOrder + "and PartNo:" + dbCheck.MaterialDesc + "and OperationNo:" + dbCheck.OperationNo;
                                responces.Add(responce);

                                isValid = false;
                            }
                        }

                        if (isValid)
                        {
                            int PrvProcessQty = 0, PrvDeliveredQty = 0, TotalProcessQty = 0, ishold = 0;
                            int isHMIFirst = 2;
                            var hmiData = db.TblSplivehmiscreen.Where(m => m.WorkOrderNo == WONoInner && m.OperationNo == OperationInner && m.PartNo == PartInner && m.IsWorkInProgress != 2).OrderByDescending(m => m.Time).ToList();

                            if (hmiData.Count > 0)
                            {
                                isHMIFirst = 0;
                            }

                            if (isHMIFirst == 0)
                            {
                                string delivString = Convert.ToString(hmiData[0].DeliveredQty);
                                int delivInt = 0;
                                int.TryParse(delivString, out delivInt);

                                string processString = Convert.ToString(hmiData[0].ProcessQty);
                                int procInt = 0;
                                int.TryParse(processString, out procInt);

                                PrvProcessQty += procInt;
                                PrvDeliveredQty += delivInt;

                                ishold = hmiData[0].IsHold;
                                ishold = ishold == 2 ? 0 : ishold;
                            }

                            TotalProcessQty = PrvProcessQty + PrvDeliveredQty;
                            int ReworkOrder = 0;
                            Tbllivehmiscreen hmidata = new Tbllivehmiscreen();
                            if (data.IsWorkOrder == 1)
                            {
                                hmidata.IsWorkOrder = 1;
                            }
                            else
                            {
                                hmidata.IsWorkOrder = 0;
                            }

                            var check = db.Tblbatchhmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner && m.IsWorkInProgress != 2).FirstOrDefault();
                            if (check == null)
                            {
                                Tblbatchhmiscreen tblbatchhmiscreen = new Tblbatchhmiscreen();
                                tblbatchhmiscreen.OperatiorId = null;
                                tblbatchhmiscreen.Status = 0;
                                tblbatchhmiscreen.Project = ddldataInner.Project;
                                tblbatchhmiscreen.PartNo = PartInner;
                                tblbatchhmiscreen.OperationNo = OperationInner;
                                tblbatchhmiscreen.WorkOrderNo = WONoInner;
                                tblbatchhmiscreen.TargetQty = Convert.ToInt32(ddldataInner.TargetQty);
                                tblbatchhmiscreen.CorrectedDate = getDate;
                                tblbatchhmiscreen.ProdFai = Type;
                                tblbatchhmiscreen.PestartTime = DateTime.Now;
                                tblbatchhmiscreen.DdlwokrCentre = ddldataInner.WorkCenter;
                                tblbatchhmiscreen.AutoBatchNumber = data.UniqueBatchNo;
                                tblbatchhmiscreen.MachineId = data.MachineId;
                                tblbatchhmiscreen.IsWorkOrder = hmidata.IsWorkOrder;
                                tblbatchhmiscreen.IsChecked = 0;
                                db.Tblbatchhmiscreen.Add(tblbatchhmiscreen);
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = tblbatchhmiscreen.AutoBatchNumber;
                                responce.message = "WorkOrder:" + WONoInner + "Operation No:" + OperationInner + "Part No:" + PartInner + "  Added Successfully";
                                responces.Add(responce);

                                if (ddldataInner != null)
                                {
                                    ddldataInner.IsWoselected = 1;
                                    ddldataInner.Bhmiid = tblbatchhmiscreen.Bhmiid;
                                    db.SaveChanges();
                                    obj.isStatus = true;
                                }
                            }
                            else
                            {
                                obj.isStatus = false;
                                responce.message = "WorkOrder:" + WONoInner + "Operation No:" + OperationInner + "Part No:" + PartInner + "  Work Order Already Exist";
                                responces.Add(responce);
                            }
                        }
                    }
                    else
                    {
                        obj.isStatus = false;
                        responce.message = "WorkOrder:" + ddldataInner.WorkOrder + "Operation No:" + ddldataInner.OperationNo + "Part No:" + ddldataInner.PartName + " Work Order Already Selected";
                        responces.Add(responce);

                    }
                }
                obj.responce = responces;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <summary>
        /// After Selecting WO from DDL 
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public AddWOCommonResponse AddBatchDetails(BatchDetails data)
        {
            AddWOCommonResponse obj = new AddWOCommonResponse();
            try
            {
                string getDate = DateTime.Now.ToString("yyyy-MM-dd");
                string[] aary = data.DdlIds.Split(',');
                List<int> intArry = aary.ToList().Select(int.Parse).ToList();
                var DDLData = db.Tblddl.Where(m => intArry.Contains(m.Ddlid)).ToList();
                List<Responce> responces = new List<Responce>();
                foreach (var item in DDLData)
                {
                    Responce responce = new Responce();
                    int DDLID = item.Ddlid;
                    var ddldataInner = db.Tblddl.Where(m => m.IsCompleted == 0 && m.Ddlid == DDLID && m.IsWoselected == 0).FirstOrDefault();
                    bool isValid = true;
                    if (ddldataInner != null)
                    {
                        string SplitWOInner = ddldataInner.SplitWo;
                        string WONoInner = ddldataInner.WorkOrder;
                        string PartInner = ddldataInner.PartName;
                        string OperationInner = ddldataInner.OperationNo;
                        string Type = ddldataInner.Type;

                        var HMICheck = db.TblSplivehmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner).FirstOrDefault();
                        if (HMICheck != null)
                        {
                            if (HMICheck.Date != null && HMICheck.Time == null)
                            {
                                responce.message = "WorkOrderNo:" + WONoInner + "and PartNo:" + PartInner + "and OperationNo:" + OperationInner + " Work Order Already Started";
                                responces.Add(responce);
                                isValid = false;
                            }
                        }

                        var HMIBatchCheck = db.Tblbatchhmiscreen.Where(m => m.OperationNo == OperationInner && m.WorkOrderNo == WONoInner && m.PartNo == PartInner && m.IsChecked != 2).FirstOrDefault();
                        if (HMIBatchCheck != null)
                        {
                            if (HMIBatchCheck.Date == null)
                            {
                                obj.isStatus = false;
                                responce.message = "WorkOrderNo:" + WONoInner + "and PartNo:" + PartInner + "and OperationNo:" + OperationInner + "  Work Order Already Selected";
                                responces.Add(responce);
                                isValid = false;
                            }
                        }

                        //var HMISequenceCheck = db.TblSplivehmiscreen.Where(m => m.WorkOrderNo == WONoInner && m.OperationNo == OperationInner && m.PartNo == PartInner).Select(m => m.OperationNo).FirstOrDefault();
                        //int OperationNo = Convert.ToInt32(HMISequenceCheck);
                        //if (HMISequenceCheck != null)
                        //{
                        //    var dbCheck = db.Tblddl.Where(m => m.WorkOrder == WONoInner && Convert.ToInt32(m.OperationNo) > OperationNo).FirstOrDefault();
                        //    if (dbCheck != null)
                        //    {
                        //        obj.isStatus = false;
                        //        responce.message = "Select WorkOrderNo:" + dbCheck.WorkOrder + "and PartNo:" + dbCheck.MaterialDesc + "and OperationNo:" + dbCheck.OperationNo;
                        //        responces.Add(responce);

                        //        isValid = false;
                        //    }
                        //}

                        var DdlSequenceCheck = (from wf in db.Tblddl
                                                where wf.WorkOrder == WONoInner && wf.MaterialDesc == PartInner && wf.IsWoselected == 0 && wf.IsCompleted == 0 && wf.IsDeleted == 0

                                                select new
                                                {
                                                    OperationNo = wf.OperationNo,
                                                    WorkOrder = wf.WorkOrder,
                                                    MaterialDesc = wf.MaterialDesc,
                                                    WorkCenter = wf.WorkCenter
                                                }).OrderBy(m => int.Parse(m.OperationNo)).ToList().FirstOrDefault();

                        if (DdlSequenceCheck != null)
                        {
                            if (Convert.ToInt32(DdlSequenceCheck.OperationNo) < Convert.ToInt32(OperationInner))
                            {
                                obj.isStatus = false;
                                responce.message = "Select WorkOrderNo:" + DdlSequenceCheck.WorkOrder + "and PartNo:" + DdlSequenceCheck.MaterialDesc + "and OperationNo:" + DdlSequenceCheck.OperationNo;
                                responces.Add(responce);
                                isValid = false;
                            }
                        }

                        if (isValid)
                        {
                            int PrvProcessQty = 0, PrvDeliveredQty = 0, TotalProcessQty = 0, ishold = 0;
                            int isHMIFirst = 2;
                            var hmiData = db.TblSplivehmiscreen.Where(m => m.WorkOrderNo == WONoInner && m.OperationNo == OperationInner && m.PartNo == PartInner && m.IsWorkInProgress != 2).OrderByDescending(m => m.Time).ToList();

                            if (hmiData.Count > 0)
                            {
                                isHMIFirst = 0;
                            }

                            if (isHMIFirst == 0)
                            {
                                string delivString = Convert.ToString(hmiData[0].DeliveredQty);
                                int delivInt = 0;
                                int.TryParse(delivString, out delivInt);

                                string processString = Convert.ToString(hmiData[0].ProcessQty);
                                int procInt = 0;
                                int.TryParse(processString, out procInt);

                                PrvProcessQty += procInt;
                                PrvDeliveredQty += delivInt;

                                ishold = hmiData[0].IsHold;
                                ishold = ishold == 2 ? 0 : ishold;
                            }

                            TotalProcessQty = PrvProcessQty + PrvDeliveredQty;
                            int ReworkOrder = 0;
                            Tbllivehmiscreen hmidata = new Tbllivehmiscreen();
                            if (data.IsWorkOrder == 1)
                            {
                                hmidata.IsWorkOrder = 1;
                            }
                            else
                            {
                                hmidata.IsWorkOrder = 0;
                            }

                            var check = db.Tblbatchhmiscreen.Where(m => m.OperationNo == OperationInner && m.AutoBatchNumber == data.UniqueBatchNo && m.WorkOrderNo == WONoInner && m.PartNo == PartInner && m.IsWorkInProgress != 2 && m.IsChecked == 2).FirstOrDefault();
                            if (check == null)
                            {
                                Tblbatchhmiscreen tblbatchhmiscreen = new Tblbatchhmiscreen();
                                tblbatchhmiscreen.Status = 0;
                                tblbatchhmiscreen.Project = ddldataInner.Project;
                                tblbatchhmiscreen.PartNo = PartInner;
                                tblbatchhmiscreen.OperationNo = OperationInner;
                                tblbatchhmiscreen.WorkOrderNo = WONoInner;
                                tblbatchhmiscreen.TargetQty = Convert.ToInt32(ddldataInner.TargetQty);
                                tblbatchhmiscreen.CorrectedDate = getDate;
                                tblbatchhmiscreen.ProdFai = Type;
                                tblbatchhmiscreen.PestartTime = DateTime.Now;
                                tblbatchhmiscreen.DdlwokrCentre = ddldataInner.WorkCenter;
                                tblbatchhmiscreen.AutoBatchNumber = data.UniqueBatchNo;
                                tblbatchhmiscreen.MachineId = data.MachineId;
                                tblbatchhmiscreen.IsWorkOrder = hmidata.IsWorkOrder;
                                tblbatchhmiscreen.IsChecked = 1;
                                tblbatchhmiscreen.PcpNo = ddldataInner.PcpNo;
                                db.Tblbatchhmiscreen.Add(tblbatchhmiscreen);
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = tblbatchhmiscreen.AutoBatchNumber;
                                responce.message = "WorkOrder:" + WONoInner + "Operation No:" + OperationInner + "Part No:" + PartInner + "  Added Successfully";
                                responces.Add(responce);

                                if (ddldataInner != null)
                                {
                                    ddldataInner.IsWoselected = 1;
                                    ddldataInner.Bhmiid = tblbatchhmiscreen.Bhmiid;
                                    db.SaveChanges();
                                    obj.isStatus = true;
                                }
                            }
                            else
                            {
                                obj.isStatus = false;
                                responce.message = "WorkOrder:" + WONoInner + "Operation No:" + OperationInner + "Part No:" + PartInner + "  Work Order Already Exist";
                                responces.Add(responce);
                            }
                        }
                    }

                }
                obj.responce = responces;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Idle Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>  
        /// 

        /// <summary>
        /// Get Idle Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public HoldIdleCodeResponse GetIdleCodes(BatchProcessingEntity.IdleCodeDetails data)
        {
            HoldIdleCodeResponse obj = new HoldIdleCodeResponse();
            try
            {
                bool idle = false;
                var machineDispName = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID).Select(m => m.MachineDispName).FirstOrDefault();

                if (data.IsTrue == false)
                {
                    var idleclick = db.TblSplivehmiscreen.Where(m => m.MachineId == data.MachineID && m.IsBatchStart == 1 && m.IsBatchFinish == 0 && m.IsPartialFinish == 0 && m.IsActivityFinish == 0).ToList();
                    if (idleclick.Count != 0)
                    {
                        foreach (var breakdownrow in idleclick)
                        {
                            if (breakdownrow.IsBreakdownClicked == 1 || breakdownrow.IsBreakdownClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "BreakDown is already Running so you cannot able to start Idle, first End BreakDown then start Idle";
                                idle = true;
                            }
                            else if (breakdownrow.IspmClicked == 1 || breakdownrow.IspmClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Planned Maintainance is already Running so you cannot able to start Idle, first End Planned Maintainance then start Idle";
                                idle = true;
                            }
                            else if (breakdownrow.IsHoldClicked == 1)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Hold is already Running so you cannot able to start Idle, first End Hold then start Idle";
                                idle = true;
                            }
                            else if (breakdownrow.IsIdleClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Overall Idle is already Running so you cannot able to Individual Idle, first End Overall Idle then start Individual Idle";
                                idle = true;
                            }
                        }

                    }

                    if (idle == true)
                    {

                    }
                    else
                    {


                        if (data.LossCodeID != 0)
                        {
                            var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.LossCodeID).FirstOrDefault();
                            //int level = lossdata.LossCodesLevel;
                            string losscode = lossdata.LossCode;
                            if (data.Level == 1)
                            {
                                var level2Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == data.LossCodeID && wf.LossCodesLevel == 2 && wf.LossCodesLevel2Id == null && wf.MessageType != "BREAKDOWN"
                                                  select new
                                                  {
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      MessageType = wf.MessageType,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id
                                                  }).ToList();
                                if (level2Data.Count == 0)
                                {
                                    obj.isStatus = false;
                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                                    obj.LossId = data.LossCodeID;
                                    obj.Level = data.Level;
                                    obj.MachineDisplayName = machineDispName;
                                }
                                else
                                {
                                    obj.isStatus = true;
                                    obj.response = level2Data;
                                    obj.Level = data.Level + 1;
                                    obj.LossId = data.LossCodeID;
                                    obj.MachineDisplayName = machineDispName;
                                }
                            }
                            else if (data.Level == 2)
                            {
                                var level3Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == data.LossCodeID && wf.LossCodesLevel == 3 && wf.MessageType != "BREAKDOWN"
                                                  select new
                                                  {
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      MessageType = wf.MessageType,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id,

                                                  }).ToList();
                                int prevLevelId = Convert.ToInt32(lossdata.LossCodesLevel1Id);
                                if (level3Data.Count == 0)
                                {
                                    obj.isStatus = false;
                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                                    obj.LossId = data.LossCodeID;
                                    obj.Level = data.Level;
                                    obj.MachineDisplayName = machineDispName;
                                }
                                else
                                {
                                    obj.isStatus = true;
                                    obj.response = level3Data;
                                    obj.Level = data.Level + 1;
                                    obj.LossId = data.LossCodeID;
                                    obj.MachineDisplayName = machineDispName;
                                }
                            }
                            else if (data.Level == 3)
                            {
                                int prevLevelId = Convert.ToInt32(lossdata.LossCodesLevel2Id);
                                var level2Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == prevLevelId && wf.LossCodesLevel == 3
                                                  select new
                                                  {
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      MessageType = wf.MessageType,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id,

                                                  }).ToList();
                                obj.isStatus = true;
                                obj.response = level2Data;
                                obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                                obj.LossId = data.LossCodeID;
                                obj.Level = 3;
                                obj.MachineDisplayName = machineDispName;
                            }
                        }
                        else
                        {
                            var level1Data = (from wf in db.Tbllossescodes
                                              where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType != "NoCode" && wf.MessageType != "BREAKDOWN" && wf.MessageType != "PM"
                                              select new
                                              {
                                                  LossCodeId = wf.LossCodeId,
                                                  LossCode = wf.LossCode,
                                                  LossCodeDesc = wf.LossCodeDesc,
                                                  MessageType = wf.MessageType,
                                                  LossCodesLevel = wf.LossCodesLevel,
                                                  LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                  LossCodesLevel2Id = wf.LossCodesLevel2Id,

                                              }).ToList();

                            if (level1Data.Count != 0)
                            {
                                obj.isStatus = true;
                                obj.response = level1Data;
                                obj.Level = data.Level + 1;
                                obj.MachineDisplayName = machineDispName;
                            }
                        }
                    }
                }
                else
                {
                    var lossdet = db.TblSplivelossofentry.Where(m => m.MachineId == data.MachineID && m.DoneWithRow == 0).OrderByDescending(m => m.SplossId).FirstOrDefault();
                    var losscode = db.Tbllossescodes.Where(m => m.LossCodeId == lossdet.MessageCodeId).FirstOrDefault();
                    obj.LossId = losscode.LossCodeId;
                    obj.losscode = losscode.LossCode;
                    obj.StartTime = Convert.ToString(lossdet.StartDateTime);
                    obj.response = "End Idle";
                    obj.MachineDisplayName = machineDispName;
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
        /// Get Hold Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <summary>
        /// Get Hold Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public HoldIdleCodeResponse GetHoldCodes(HoldCodeDetails data)
        {
            HoldIdleCodeResponse obj = new HoldIdleCodeResponse();
            try
            {
                bool idle = false;
                var machineDispName = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID).Select(m => m.MachineDispName).FirstOrDefault();
                if (data.IsTrue == false)
                {
                    var holdclick = db.TblSplivehmiscreen.Where(m => m.MachineId == data.MachineID && m.IsBatchStart == 1 && m.IsBatchFinish == 0 && m.IsPartialFinish == 0 && m.IsActivityFinish == 0).ToList();
                    if (holdclick.Count != 0)
                    {
                        foreach (var breakdownrow in holdclick)
                        {
                            if (breakdownrow.IsBreakdownClicked == 1 || breakdownrow.IsBreakdownClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "BreakDown is already Running so you cannot able to start Hold first End BreakDown then start Hold";
                                idle = true;
                            }
                            else if (breakdownrow.IspmClicked == 1 || breakdownrow.IspmClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Planned Maintainance is already Running so you cannot able to start Hold first End Planned Maintainance then start Hold";
                                idle = true;
                            }
                            else if (breakdownrow.IsIdleClicked == 1 || breakdownrow.IsIdleClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Idle is already Running so you cannot able to start Hold first End Idle then start Hold";
                                idle = true;
                            }
                        }

                    }

                    if (idle == true)
                    {

                    }
                    else
                    {
                        if (data.HoldCodeID != 0)
                        {
                            var check = db.Tblholdcodes.Where(m => m.HoldCodeId == data.HoldCodeID).FirstOrDefault();
                            //var level = check.HoldCodesLevel;
                            string lossCodes = check.HoldCode;
                            if (data.Level == 1)
                            {
                                var level2Data = (from wf in db.Tblholdcodes
                                                  where wf.IsDeleted == 0 && wf.HoldCodesLevel1Id == data.HoldCodeID && wf.HoldCodesLevel == 2 && wf.HoldCodesLevel2Id == null
                                                  select new
                                                  {
                                                      HoldCode = wf.HoldCode,
                                                      HoldCodeDesc = wf.HoldCodeDesc,
                                                      HoldCodeId = wf.HoldCodeId,
                                                      HoldCodesLevel = wf.HoldCodesLevel,
                                                      HoldCodesLevel1Id = wf.HoldCodesLevel1Id,
                                                      HoldCodesLevel2Id = wf.HoldCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }
                                                  ).ToList();

                                if (level2Data.Count == 0)
                                {

                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + lossCodes + " as reason.";
                                    obj.LossId = data.HoldCodeID;
                                    obj.Level = data.Level;
                                    obj.isStatus = false;
                                    obj.MachineDisplayName = machineDispName;
                                }
                                else
                                {
                                    obj.isStatus = true;
                                    obj.response = level2Data;
                                    obj.Level = data.Level + 1;
                                    obj.LossId = data.HoldCodeID;
                                    obj.MachineDisplayName = machineDispName;
                                }
                            }
                            else if (data.Level == 2)
                            {
                                var level3Data = (from wf in db.Tblholdcodes
                                                  where wf.IsDeleted == 0 && wf.HoldCodesLevel2Id == data.HoldCodeID && wf.HoldCodesLevel == 3
                                                  select new
                                                  {
                                                      HoldCode = wf.HoldCode,
                                                      HoldCodeDesc = wf.HoldCodeDesc,
                                                      HoldCodeId = wf.HoldCodeId,
                                                      HoldCodesLevel = wf.HoldCodesLevel,
                                                      HoldCodesLevel1Id = wf.HoldCodesLevel1Id,
                                                      HoldCodesLevel2Id = wf.HoldCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }).ToList();
                                int prevLevelId = Convert.ToInt32(check.HoldCodesLevel1Id);

                                if (level3Data.Count == 0)
                                {

                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + lossCodes + " as reason.";
                                    obj.LossId = data.HoldCodeID;
                                    obj.Level = data.Level;
                                    obj.isStatus = false;
                                    obj.MachineDisplayName = machineDispName;
                                }
                                else
                                {
                                    obj.isStatus = true;
                                    obj.response = level3Data;
                                    obj.Level = data.Level + 1;
                                    obj.LossId = data.HoldCodeID;
                                    obj.MachineDisplayName = machineDispName;
                                }
                            }
                            else if (data.Level == 3)
                            {
                                int prevLevelId = Convert.ToInt32(check.HoldCodesLevel2Id);
                                int FirstLevelID = Convert.ToInt32(check.HoldCodesLevel1Id);
                                var level2Data = (from wf in db.Tblholdcodes
                                                  where wf.IsDeleted == 0 && wf.HoldCodesLevel2Id == prevLevelId && wf.HoldCodesLevel == 3
                                                  select new
                                                  {
                                                      HoldCode = wf.HoldCode,
                                                      HoldCodeDesc = wf.HoldCodeDesc,
                                                      HoldCodeId = wf.HoldCodeId,
                                                      HoldCodesLevel = wf.HoldCodesLevel,
                                                      HoldCodesLevel1Id = wf.HoldCodesLevel1Id,
                                                      HoldCodesLevel2Id = wf.HoldCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }).ToList();
                                obj.isStatus = true;
                                obj.response = level2Data;
                                obj.ItsLastLevel = "No Further Levels . Do you want to set " + lossCodes + " as reason.";
                                obj.LossId = data.HoldCodeID;
                                obj.Level = 3;
                                obj.MachineDisplayName = machineDispName;
                            }
                        }
                        else
                        {
                            var check = (from wf in db.Tblholdcodes
                                         where wf.HoldCodesLevel == 1 && wf.IsDeleted == 0
                                         select new
                                         {
                                             HoldCode = wf.HoldCode,
                                             HoldCodeDesc = wf.HoldCodeDesc,
                                             HoldCodeId = wf.HoldCodeId,
                                             HoldCodesLevel = wf.HoldCodesLevel,
                                             HoldCodesLevel1Id = wf.HoldCodesLevel1Id,
                                             HoldCodesLevel2Id = wf.HoldCodesLevel2Id,
                                             MessageType = wf.MessageType
                                         }).ToList();
                            if (check.Count != 0)
                            {
                                obj.isStatus = true;
                                obj.response = check;
                                obj.MachineDisplayName = machineDispName;
                                obj.Level = data.Level + 1;
                                obj.MachineDisplayName = machineDispName;
                            }
                            else
                            {
                                obj.isStatus = false;
                                obj.response = "No Items Found";
                            }
                        }
                    }



                }
                else
                {
                    obj.isStatus = true;
                    obj.response = "Do You Want to Release";
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
        /// Get Break Down Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <summary>
        /// Get Break Down Codes
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public BreakDownCodeResponse GetBreakDownCodes(AddBreakDownCodes data)
        {
            BreakDownCodeResponse obj = new BreakDownCodeResponse();
            try
            {
                bool breakdown = false;
                var machinedispname = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineId).Select(m => m.MachineDispName).FirstOrDefault();
                if (data.IsTrue == false)
                {
                    var breakdownclick = db.TblSplivehmiscreen.Where(m => m.MachineId == data.MachineId && m.IsBatchStart == 1 && m.IsBatchFinish == 0 && m.IsPartialFinish == 0 && m.IsActivityFinish == 0).ToList();
                    if (breakdownclick.Count != 0)
                    {
                        foreach (var breakdownrow in breakdownclick)
                        {
                            if (breakdownrow.IsIdleClicked == 1 || breakdownrow.IsIdleClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Idle is already Running so you cannot able to start BreakDown, first End Idle then start BreakDown";
                                breakdown = true;
                            }
                            else if (breakdownrow.IspmClicked == 1 || breakdownrow.IspmClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Planned Maintainance is already Running so you cannot able to start BreakDown, first End Planned Maintainance then start BreakDown";
                                breakdown = true;
                            }
                            else if (breakdownrow.IsHoldClicked == 1)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Hold is already Running so you cannot able to start Overall BreakDown, first End Hold then start Overall BreakDown";
                                breakdown = true;
                            }
                            else if (breakdownrow.IsBreakdownClicked == 2)
                            {
                                obj.isStatus = false;
                                obj.errorMsg = "Overall Breakdown is already Running so you cannot able to start Individual BreakDown, first End Overall Breakdown then start Individual BreakDown";
                                breakdown = true;
                            }
                        }

                    }
                    if (breakdown == true)
                    {

                    }
                    else
                    {


                        var modedet = GetMode(data.MachineId);

                        var curMode = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineId && m.DoneWithRow == 0).OrderByDescending(m => m.SpbreakdownId).Take(1).ToList();
                        int currentId = 0;
                        foreach (var item in curMode)
                        {
                            currentId = item.SpbreakdownId;
                            var loss = db.Tbllossescodes.Where(m => m.LossCodeId == item.BreakDownCode).FirstOrDefault();
                            string mode = loss.MessageType;
                            if (mode == "PM")
                            {
                                obj.response = "Machine is in Maintenance , cannot change mode to Breakdown";
                                obj.isStatus = false;
                            }
                            else if (mode != "BREAKDOWN")
                            {
                                TblSpbreakdown tblSpbreakdown = db.TblSpbreakdown.Where(m => m.SpbreakdownId == currentId).FirstOrDefault();
                                tblSpbreakdown.EndTime = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                if (tblSpbreakdown != null)
                                {
                                    tblSpbreakdown.StartTime = tblSpbreakdown.StartTime;
                                    tblSpbreakdown.EndTime = tblSpbreakdown.EndTime;
                                    tblSpbreakdown.MachineId = tblSpbreakdown.MachineId;
                                    tblSpbreakdown.CorrectedDate = tblSpbreakdown.CorrectedDate;
                                    tblSpbreakdown.BreakDownCode = tblSpbreakdown.BreakDownCode;
                                    tblSpbreakdown.Shift = tblSpbreakdown.Shift;
                                    tblSpbreakdown.MessageDesc = tblSpbreakdown.MessageDesc;
                                    tblSpbreakdown.MessageCode = tblSpbreakdown.MessageCode;
                                    tblSpbreakdown.DoneWithRow = tblSpbreakdown.DoneWithRow;
                                    db.SaveChanges();
                                    obj.isStatus = true;
                                    obj.response = "Added Successfully";
                                    break;
                                }
                            }
                        }

                        var breakdownToView = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineId && m.DoneWithRow == 0).OrderByDescending(m => m.SpbreakdownId).FirstOrDefault();
                        if (breakdownToView != null) //implies brekdown is running
                        {
                            if (breakdownToView.DoneWithRow == 0)
                            {
                                int breakdowncode = Convert.ToInt32(breakdownToView.BreakDownCode);
                                var DataToView = db.Tbllossescodes.Where(m => m.LossCodeId == breakdowncode).ToList();
                                obj.Level = DataToView[0].LossCodesLevel;
                                obj.BreakdownCode = DataToView[0].LossCode;
                                obj.BreakdownId = DataToView[0].LossCodeId;
                                obj.BreakdownStartTime = Convert.ToString(breakdownToView.StartTime);
                                obj.response = DataToView;
                                obj.isStatus = true;
                            }
                        }

                        if (data.LossCodeID != 0)
                        {
                            var breakdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.LossCodeID).FirstOrDefault();
                            //int level = breakdata.LossCodesLevel;
                            string breakdowncode = breakdata.LossCode;
                            if (data.Level == 1)
                            {
                                var level2Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel1Id == data.LossCodeID && wf.LossCodesLevel2Id == null && wf.MessageType == "BREAKDOWN"
                                                  select new
                                                  {
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }).ToList();
                                if (level2Data.Count == 0)
                                {

                                    obj.isStatus = false;
                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + breakdowncode + " as reason.";
                                    obj.BreakdownId = data.LossCodeID;
                                    obj.Level = data.Level;
                                }
                                else
                                {
                                    obj.isStatus = true;
                                    obj.response = level2Data;
                                    obj.Level = data.Level + 1;
                                    obj.BreakdownId = data.LossCodeID;
                                }
                            }
                            else if (data.Level == 2)
                            {
                                var level3Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == data.LossCodeID && wf.MessageType == "BREAKDOWN"
                                                  select new
                                                  {
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }).ToList();
                                int prevLevelId = Convert.ToInt32(breakdata.LossCodesLevel1Id);
                                if (level3Data.Count == 0)
                                {


                                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + breakdowncode + " as reason.";
                                    obj.BreakdownId = data.LossCodeID;
                                    obj.Level = data.Level;
                                }
                                else
                                {
                                    obj.Level = data.Level + 1;
                                    obj.BreakdownId = data.LossCodeID;
                                    obj.response = level3Data;
                                }
                            }
                            else if (data.Level == 3)
                            {
                                int prevLevelId = Convert.ToInt32(breakdata.LossCodesLevel2Id);
                                int FirstLevelID = Convert.ToInt32(breakdata.LossCodesLevel1Id);
                                var level2Data = (from wf in db.Tbllossescodes
                                                  where wf.IsDeleted == 0 && wf.LossCodesLevel2Id == prevLevelId && wf.MessageType == "BREAKDOWN"
                                                  select new
                                                  {
                                                      LossCode = wf.LossCode,
                                                      LossCodeDesc = wf.LossCodeDesc,
                                                      LossCodeId = wf.LossCodeId,
                                                      LossCodesLevel = wf.LossCodesLevel,
                                                      LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                      LossCodesLevel2Id = wf.LossCodesLevel2Id,
                                                      MessageType = wf.MessageType
                                                  }).ToList();

                                obj.ItsLastLevel = "No Further Levels . Do you want to set " + breakdowncode + " as reason.";
                                obj.BreakdownId = data.LossCodeID;
                                obj.Level = 3;
                                obj.response = level2Data;
                            }
                        }
                        else
                        {
                            var level1Data = (from wf in db.Tbllossescodes
                                              where wf.IsDeleted == 0 && wf.LossCodesLevel == 1 && wf.MessageType == "BREAKDOWN" && wf.LossCode != "9999"
                                              select new
                                              {
                                                  LossCode = wf.LossCode,
                                                  LossCodeDesc = wf.LossCodeDesc,
                                                  LossCodeId = wf.LossCodeId,
                                                  LossCodesLevel = wf.LossCodesLevel,
                                                  LossCodesLevel1Id = wf.LossCodesLevel1Id,
                                                  LossCodesLevel2Id = wf.LossCodesLevel2Id,
                                                  MessageType = wf.MessageType
                                              }).ToList();
                            obj.isStatus = true;
                            obj.response = level1Data;
                            obj.Level = data.Level + 1;
                            obj.MachineDisplayName = machinedispname;

                        }
                    }
                }
                else
                {
                    var breakdownToView = db.TblSpbreakdown.Where(m => m.MachineId == data.MachineId && m.DoneWithRow == 0).OrderByDescending(m => m.SpbreakdownId).FirstOrDefault();
                    int losscodeid = (int)breakdownToView.BreakDownCode;
                    obj.BreakdownId = losscodeid;
                    var losscode = db.Tbllossescodes.Where(m => m.LossCodeId == losscodeid).FirstOrDefault();
                    obj.BreakdownCode = losscode.LossCode;
                    obj.BreakdownStartTime = Convert.ToString(breakdownToView.StartTime);
                    obj.response = "End BreakDown";
                    obj.MachineDisplayName = machinedispname;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        public string GetMode(int MachineID)
        {
            string res = "";
            var modedata = db.Tbllivemodedb.Where(m => m.MachineId == MachineID && m.IsCompleted == 0).OrderByDescending(m => m.StartTime).FirstOrDefault();
            if (modedata != null)
            {
                var mode = modedata.Mode;
                res = mode;
            }

            return res;
        }

        //Monika
        //public CommonResponseWithMachinedesscName GetStartedBatch(int MachineID)
        //{
        //    CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
        //    try
        //    {
        //        var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == MachineID && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();
        //        var shift = GetDateShift();


        //        List<GetCreatedBatches> wolist = new List<GetCreatedBatches>();
        //        var BatchDet = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 0 && m.IsChecked !=2).ToList();
        //        if (BatchDet.Count > 0)
        //        {
        //            var createdBatchDet = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 0 && m.IsBatchFinish == 0 && m.IsActivityFinish == 0 && m.IsChecked !=2).Select(m=>m.AutoBatchNumber).Distinct().ToList();
        //            foreach (var batchrow in createdBatchDet)
        //            {
        //                GetCreatedBatches get = new GetCreatedBatches();
        //                string batchno = Convert.ToString(batchrow);
        //                get.autoBatchNumber = batchno;
        //                wolist.Add(get);

        //                var getdata2 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == batchno && m.IsGenericClicked == 1).FirstOrDefault();
        //                if (getdata2 != null)
        //                {
        //                    comres.isGeneric = true;
        //                }
        //            }
        //        }
        //        var BatchDetWhenStart = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 1 && m.IsChecked != 2).ToList();
        //        if (BatchDetWhenStart.Count > 0)
        //        {
        //            var batchDet = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 1 && m.IsBatchFinish == 0 && (m.IsPatialFinish == 1 || m.IsActivityFinish == 1 && m.IsChecked != 2)).Select(m => m.AutoBatchNumber).Distinct().ToList();
        //            foreach (var batchrow in batchDet)
        //            {
        //                GetCreatedBatches get = new GetCreatedBatches();
        //                string batchno = Convert.ToString(batchrow);
        //                get.autoBatchNumber = batchrow;
        //                wolist.Add(get);

        //                var getdata2 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == batchno && m.IsGenericClicked == 1).FirstOrDefault();
        //                if (getdata2 != null)
        //                {
        //                    comres.isGeneric = true;
        //                }
        //            }
        //        }
        //        comres.isTure = true;
        //        comres.response = wolist;
        //        comres.MacDispName = machinedet;
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //}

        public CommonResponseWithMachinedesscName GetStartedBatch(int MachineID)
        {
            CommonResponseWithMachinedesscName comres = new CommonResponseWithMachinedesscName();
            try
            {
                var machinedet = db.Tblmachinedetails.Where(m => m.MachineId == MachineID && m.IsDeleted == 0).Select(m => m.MachineDispName).FirstOrDefault();
                var shift = GetDateShift();


                List<GetCreatedBatches> wolist = new List<GetCreatedBatches>();
                var BatchDet = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 0 && m.IsChecked != 2).ToList();
                if (BatchDet.Count > 0)
                {
                    var createdBatchDet = db.Tblbatchhmiscreen.Where(m => m.MachineId == MachineID && m.IsBatchStart == 0 && m.IsChecked != 2).Select(m => m.AutoBatchNumber).Distinct().ToList();
                    foreach (var batchrow in createdBatchDet)
                    {
                        GetCreatedBatches get = new GetCreatedBatches();
                        string batchno = Convert.ToString(batchrow);
                        get.autoBatchNumber = batchno;
                        wolist.Add(get);

                        var getdata2 = db.Tblbatchhmiscreen.Where(m => m.AutoBatchNumber == batchno && m.IsGenericClicked == 1).FirstOrDefault();
                        if (getdata2 != null)
                        {
                            comres.isGeneric = true;
                        }
                    }
                }
                comres.isTure = true;
                comres.response = wolist;
                comres.MacDispName = machinedet;
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        /// <summary>
        /// Add Idle to List of Batch Number
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponsewithEror AddBatchNoListToIdle(AddIdleCodesList data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            try
            {
                string[] ids = data.batchNo.Split(',');
                foreach (var item in ids)
                {
                    //int BatchNo = Convert.ToInt32(item);
                    string CorrectedDate = null;
                    Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                    TimeSpan Start = StartTime.StartTime;
                    if (Start <= DateTime.Now.TimeOfDay)
                    {
                        CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                    }

                    DateTime Time = DateTime.Now;
                    TimeSpan Tm = new TimeSpan(Time.Hour, Time.Minute, Time.Second);
                    var ShiftDetails = db.TblshiftMstr.Where(m => m.StartTime <= Tm && m.EndTime >= Tm).FirstOrDefault();
                    string Shift = "C";
                    if (ShiftDetails != null)
                    {
                        Shift = ShiftDetails.ShiftName;
                    }

                    if (data.lossCodeId != 0 && data.status == "Start")
                    {

                        var tblbreakdown = db.TblSplivemodedb.Where(m => m.BatchNumber == item && m.IsCompleted == 0 && m.Mode == "IDLE").FirstOrDefault();
                        if (tblbreakdown == null)
                        {

                            var breakdata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeId == data.lossCodeId).FirstOrDefault();
                            string msgCode = breakdata.LossCode;
                            string msgDesc = breakdata.LossCodeDesc;

                            TblSplivelossofentry tblSplivelossofentry = new TblSplivelossofentry();
                            if (data.lossCodeId != 0)
                            {
                                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.lossCodeId).FirstOrDefault();
                                tblSplivelossofentry.MessageCodeId = data.lossCodeId;
                                tblSplivelossofentry.MessageDesc = lossdata.LossCodeDesc;
                                tblSplivelossofentry.MessageCode = lossdata.LossCode;
                            }
                            else
                            {
                                int MessageCodeID = Convert.ToInt32(tblSplivelossofentry.MessageCodeId);
                                tblSplivelossofentry.MessageCodeId = Convert.ToInt32(tblSplivelossofentry.MessageCodeId);
                                var a = db.MessageCodeMaster.Where(m => m.MessageCodeId == MessageCodeID).FirstOrDefault();
                                tblSplivelossofentry.MessageDesc = a.MessageDescription.ToString();
                                tblSplivelossofentry.MessageCode = a.MessageCode.ToString();
                            }

                            tblSplivelossofentry.StartDateTime = DateTime.Now;
                            tblSplivelossofentry.EntryTime = DateTime.Now;
                            tblSplivelossofentry.EndDateTime = DateTime.Now;
                            tblSplivelossofentry.CorrectedDate = CorrectedDate;
                            tblSplivelossofentry.MachineId = data.machineId;
                            tblSplivelossofentry.Shift = Shift;
                            tblSplivelossofentry.MessageCodeId = tblSplivelossofentry.MessageCodeId;
                            tblSplivelossofentry.MessageDesc = tblSplivelossofentry.MessageDesc;
                            tblSplivelossofentry.MessageCode = tblSplivelossofentry.MessageCode;
                            tblSplivelossofentry.IsUpdate = 1;
                            tblSplivelossofentry.DoneWithRow = 0;
                            tblSplivelossofentry.IsStart = 1;
                            tblSplivelossofentry.IsScreen = 0;
                            tblSplivelossofentry.BatchNo = item;
                            tblSplivelossofentry.ForRefresh = 1;
                            //tblSplivelossofentry. = item;
                            db.TblSplivelossofentry.Add(tblSplivelossofentry);
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = "Yellow";

                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.machineId && m.IsCompleted == 0 && m.BatchNumber == item).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.IsCompleted = 1;
                                modedata.EndTime = DateTime.Now;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = "Updated Successfully";
                            }

                            TblSplivemodedb tblSplivemodedb = new TblSplivemodedb();
                            tblSplivemodedb.MachineId = data.machineId;
                            tblSplivemodedb.CorrectedDate = CorrectedDate;
                            tblSplivemodedb.StartTime = DateTime.Now;
                            tblSplivemodedb.ColorCode = "yellow";
                            tblSplivemodedb.InsertedOn = DateTime.Now;
                            tblSplivemodedb.IsDeleted = 0;
                            tblSplivemodedb.Mode = "IDLE";
                            tblSplivemodedb.IsIdle = 1;
                            tblSplivemodedb.BatchNumber = item;
                            tblSplivemodedb.IsCompleted = 0;
                            db.TblSplivemodedb.Add(tblSplivemodedb);
                            db.SaveChanges();

                            if (data.IndividualClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 1;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                            }
                            else if (data.OverallClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 2;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                            }

                            obj.isTure = true;
                            obj.response = "Yellow";
                        }
                        else
                        {
                            var breakdata = db.Tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeId == data.lossCodeId).FirstOrDefault();
                            string msgCode = breakdata.LossCode;
                            string msgDesc = breakdata.LossCodeDesc;

                            TblSplivelossofentry tblSplivelossofentry = new TblSplivelossofentry();
                            if (data.lossCodeId != 0)
                            {
                                var lossdata = db.Tbllossescodes.Where(m => m.LossCodeId == data.lossCodeId).FirstOrDefault();
                                tblSplivelossofentry.MessageCodeId = data.lossCodeId;
                                tblSplivelossofentry.MessageDesc = lossdata.LossCodeDesc;
                                tblSplivelossofentry.MessageCode = lossdata.LossCode;
                            }
                            else
                            {
                                int MessageCodeID = Convert.ToInt32(tblSplivelossofentry.MessageCodeId);
                                tblSplivelossofentry.MessageCodeId = Convert.ToInt32(tblSplivelossofentry.MessageCodeId);
                                var a = db.MessageCodeMaster.Where(m => m.MessageCodeId == MessageCodeID).FirstOrDefault();
                                tblSplivelossofentry.MessageDesc = a.MessageDescription.ToString();
                                tblSplivelossofentry.MessageCode = a.MessageCode.ToString();
                            }

                            tblSplivelossofentry.StartDateTime = DateTime.Now;
                            tblSplivelossofentry.EntryTime = DateTime.Now;
                            tblSplivelossofentry.EndDateTime = DateTime.Now;
                            tblSplivelossofentry.CorrectedDate = CorrectedDate;
                            tblSplivelossofentry.MachineId = data.machineId;
                            tblSplivelossofentry.Shift = Shift;
                            tblSplivelossofentry.MessageCodeId = tblSplivelossofentry.MessageCodeId;
                            tblSplivelossofentry.MessageDesc = tblSplivelossofentry.MessageDesc;
                            tblSplivelossofentry.MessageCode = tblSplivelossofentry.MessageCode;
                            tblSplivelossofentry.IsUpdate = 1;
                            tblSplivelossofentry.DoneWithRow = 0;
                            tblSplivelossofentry.IsStart = 1;
                            tblSplivelossofentry.IsScreen = 0;
                            tblSplivelossofentry.BatchNo = item;
                            tblSplivelossofentry.ForRefresh = 1;
                            //tblSplivelossofentry. = item;
                            db.TblSplivelossofentry.Add(tblSplivelossofentry);
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = "Yellow";

                            if (data.IndividualClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 1;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                            }
                            else if (data.OverallClicked == true)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 2;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }

                            }

                        }
                    }
                    else if (data.lossCodeId != 0 && data.status == "End")
                    {
                        var hmiIdle = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item && m.IsIdleClicked == 1).FirstOrDefault();
                        if (hmiIdle != null)
                        {
                            var tb = db.TblSplivelossofentry.Where(m => m.MessageCodeId == data.lossCodeId && m.MachineId == data.machineId && m.DoneWithRow == 0 && m.BatchNo == item).OrderByDescending(m => m.SplossId).FirstOrDefault();
                            if (tb != null)
                            {
                                tb.EndDateTime = DateTime.Now;
                                tb.DoneWithRow = 1;
                                tb.IsUpdate = 1;
                                tb.IsScreen = 0;
                                tb.IsStart = 0;
                                tb.ForRefresh = 0;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = "Updated Successfully";
                            }
                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.machineId && m.IsCompleted == 0 && m.BatchNumber == item).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.EndTime = DateTime.Now;
                                modedata.IsCompleted = 1;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = "Updated Successfully";
                            }

                            TblSplivemodedb tblSplivemodedb = new TblSplivemodedb();
                            tblSplivemodedb.MachineId = data.machineId;
                            tblSplivemodedb.CorrectedDate = CorrectedDate;
                            tblSplivemodedb.StartTime = DateTime.Now;
                            tblSplivemodedb.ColorCode = "green";
                            tblSplivemodedb.InsertedOn = DateTime.Now;
                            tblSplivemodedb.IsDeleted = 0;
                            tblSplivemodedb.Mode = "PowerOn";
                            tblSplivemodedb.IsCompleted = 0;
                            tblSplivemodedb.IsIdle = 2;
                            tblSplivemodedb.BatchNumber = item;
                            db.TblSplivemodedb.Add(tblSplivemodedb);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item && m.IsIdleClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "Individual idle has not been started";
                                }
                            }


                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item && (m.IsIdleClicked == 2 || m.IsIdleClicked == 1)).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "Individual idle has not been started";
                                }
                            }
                        }
                        else
                        {
                            var tb = db.TblSplivelossofentry.Where(m => m.MessageCodeId == data.lossCodeId && m.MachineId == data.machineId && m.DoneWithRow == 0 && m.BatchNo == item).OrderByDescending(m => m.SplossId).FirstOrDefault();
                            if (tb != null)
                            {
                                tb.EndDateTime = DateTime.Now;
                                tb.DoneWithRow = 1;
                                tb.IsUpdate = 1;
                                tb.IsScreen = 0;
                                tb.IsStart = 0;
                                tb.ForRefresh = 0;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = "Updated Successfully";
                            }
                            var modedata = db.TblSplivemodedb.Where(m => m.MachineId == data.machineId && m.IsCompleted == 0 && m.BatchNumber == item).OrderByDescending(m => m.StartTime).FirstOrDefault();
                            if (modedata != null)
                            {
                                modedata.EndTime = DateTime.Now;
                                modedata.IsCompleted = 1;
                                double diff = DateTime.Now.Subtract(Convert.ToDateTime(modedata.StartTime)).TotalSeconds;
                                modedata.DurationInSec = (int)diff;
                                db.SaveChanges();
                                obj.isTure = true;
                                obj.response = "Updated Successfully";
                            }

                            TblSplivemodedb tblSplivemodedb = new TblSplivemodedb();
                            tblSplivemodedb.MachineId = data.machineId;
                            tblSplivemodedb.CorrectedDate = CorrectedDate;
                            tblSplivemodedb.StartTime = DateTime.Now;
                            tblSplivemodedb.ColorCode = "green";
                            tblSplivemodedb.InsertedOn = DateTime.Now;
                            tblSplivemodedb.IsDeleted = 0;
                            tblSplivemodedb.Mode = "PowerOn";
                            tblSplivemodedb.IsCompleted = 0;
                            tblSplivemodedb.IsIdle = 2;
                            tblSplivemodedb.BatchNumber = item;
                            db.TblSplivemodedb.Add(tblSplivemodedb);
                            db.SaveChanges();

                            if (data.IndividualClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item && m.IsIdleClicked == 1).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "Individual idle has not been started";
                                }
                            }


                            else if (data.OverallClicked == false)
                            {
                                var hmi = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item && m.IsIdleClicked == 2).FirstOrDefault();
                                if (hmi != null)
                                {
                                    hmi.IsIdleClicked = 0;
                                    db.Entry(hmi).State = EntityState.Modified;
                                    db.SaveChanges();
                                }
                                else
                                {
                                    obj.isTure = false;
                                    obj.response = "Individual idle has not been started";
                                }
                            }
                        }
                        obj.isTure = true;
                        obj.response = "Blue";
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// Add List of BatchNo To HoldCode
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>  
        /// 

        public GeneralResponse AddBatchNoListToHoldCode(AddHoldCodesList data)
        {
            GeneralResponse obj = new GeneralResponse();
            try
            {
                string[] batchNos = batchNos = data.BatchNo.Split(',');

                #region
                string CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                if (DateTime.Now.Hour < 6 && DateTime.Now.Hour >= 0)
                {
                    CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                #endregion
                foreach (var item in batchNos)
                {

                    var tblhmi = (from wf in db.TblSplivehmiscreen
                                  where wf.AutoBatchNumber == item
                                  select new
                                  {
                                      WorkOrderNo = wf.WorkOrderNo,
                                      PartNo = wf.PartNo,
                                      OperationNo = wf.OperationNo,
                                      BatchHmiid = wf.BatchHmiid,
                                      Sphmiid = wf.Sphmiid
                                  }).ToList();
                    if (tblhmi.Count != 0)
                    {
                        foreach (var hmi in tblhmi)
                        {
                            var check = db.TblSplivehmiscreen.Where(m => m.Sphmiid == hmi.Sphmiid).FirstOrDefault();
                            if (check != null)
                            {
                                check.Time = DateTime.Now;
                                check.IsWorkInProgress = 0;
                                check.IsHold = 1;
                                check.Status = 2;
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = "Updated Successfully";
                                obj.color = "Yellow";

                                var hmi1 = db.TblSplivehmiscreen.Where(m => m.AutoBatchNumber == item).FirstOrDefault();
                                hmi1.IsIdleClicked = 2;
                                db.Entry(hmi1).State = EntityState.Modified;
                                db.SaveChanges();
                            }

                            var holdCodeData = db.Tblholdcodes.Where(m => m.HoldCodeId == data.HoldCodeId).FirstOrDefault();

                            TblSplivemanuallossofentry tblSplivemanuallossofentry = new TblSplivemanuallossofentry();
                            tblSplivemanuallossofentry.CorrectedDate = CorrectedDate;
                            tblSplivemanuallossofentry.Hmiid = check.Sphmiid;
                            tblSplivemanuallossofentry.MessageCodeId = data.HoldCodeId;
                            tblSplivemanuallossofentry.MessageDesc = holdCodeData.HoldCodeDesc;
                            tblSplivemanuallossofentry.MessageCode = holdCodeData.HoldCode;
                            tblSplivemanuallossofentry.Wono = hmi.WorkOrderNo;
                            tblSplivemanuallossofentry.OpNo = Convert.ToInt32(hmi.OperationNo);
                            tblSplivemanuallossofentry.PartNo = hmi.PartNo;
                            tblSplivemanuallossofentry.BatchNo = item;
                            tblSplivemanuallossofentry.StartDateTime = DateTime.Now;
                            string[] GetDateShift1 = GetDateShift();
                            tblSplivemanuallossofentry.Shift = GetDateShift1[0];
                            tblSplivemanuallossofentry.MachineId = data.MachineId;
                            db.TblSplivemanuallossofentry.Add(tblSplivemanuallossofentry);
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.response = "Added Successfully";
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

        public CommonResponse GetModeDetails(int MachineId)
        {
            CommonResponse obj = new CommonResponse();

            var modeData = (from wf in db.TblSplivemodedb
                            where wf.MachineId == MachineId && wf.IsCompleted == 0
                            select new
                            {
                                wf.Mode,
                                wf.StartTime
                            }).OrderByDescending(m => m.StartTime).FirstOrDefault();
            if (modeData != null)
            {
                obj.isTure = true;
                obj.response = modeData;
            }
            else
            {
                obj.isTure = false;
                obj.response = "No Data Found";
            }
            return obj;
        }

        public BreakDownCodeResponse GetGenericWorkCodes(GenericWODetails data)
        {
            BreakDownCodeResponse obj = new BreakDownCodeResponse();
            string CorrectedDate = null;
            Tbldaytiming StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
            TimeSpan Start = StartTime.StartTime;
            if (Start <= DateTime.Now.TimeOfDay)
            {
                CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
            }
            else
            {
                CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            }
            var machinedispname = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == data.MachineID).Select(m => m.MachineDispName).FirstOrDefault();

            var HMIData = db.TblSplivehmiscreen.Where(m => m.MachineId == data.MachineID && m.Status == 0).ToList();
            if (HMIData.Count > 1)
            {
                obj.isStatus = false;
                obj.response = "Please Click JobFinish or PartialFinish Before Starting Generic Work";
            }
            else if (HMIData.Count == 1) //This one may be our empty row or not
            {
                if (HMIData[0].Date != null)
                {
                    obj.isStatus = false;
                    obj.response = "Please Finish CurrentJob Before Starting GenericWork";
                }
            }
            var IdleToView = db.TblSpgenericworkentry.Where(m => m.MachineId == data.MachineID).OrderByDescending(m => m.SpgwentryId).FirstOrDefault();
            if (IdleToView != null) //implies idle is running
            {
                if (IdleToView.DoneWithRow == 0)
                {
                    int idlecode = IdleToView.GwcodeId;
                    var DataToView = db.Tblgenericworkcodes.Where(m => m.IsDeleted == 0 && m.GenericWorkId == idlecode).ToList();
                    obj.Level = DataToView[0].GwcodesLevel;
                    obj.BreakdownCode = DataToView[0].GenericWorkCode;
                    obj.BreakdownId = DataToView[0].GenericWorkId;
                    obj.BreakdownStartTime = Convert.ToString(IdleToView.StartDateTime);
                }
            }

            if (data.GenericWorkId != 0)
            {
                var lossdata = db.Tblgenericworkcodes.Where(m => m.GenericWorkId == data.GenericWorkId).FirstOrDefault();
                string losscode = lossdata.GenericWorkCode;
                if (data.Level == 1)
                {
                    var level2Data = (from wf in db.Tblgenericworkcodes
                                      where wf.IsDeleted == 0 && wf.GwcodesLevel1Id == data.GenericWorkId && wf.GwcodesLevel == 2 && wf.GwcodesLevel2Id == null
                                      select new
                                      {
                                          GenericWorkId = wf.GenericWorkId,
                                          GenericWorkCode = wf.GenericWorkCode,
                                          GenericWorkDesc = wf.GenericWorkDesc,
                                          GwcodesLevel = wf.GwcodesLevel,
                                          GwcodesLevel1Id = wf.GwcodesLevel1Id,
                                          GwcodesLevel2Id = wf.GwcodesLevel2Id
                                      }).ToList();
                    if (level2Data.Count == 0)
                    {
                        obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                        obj.BreakdownId = data.GenericWorkId;
                        obj.Level = data.Level;
                        obj.MachineDisplayName = machinedispname;
                    }
                    else
                    {
                        obj.Level = data.Level + 1;
                        obj.BreakdownId = data.GenericWorkId;
                        obj.MachineDisplayName = machinedispname;
                        obj.isStatus = true;
                        obj.response = level2Data;
                    }
                }
                else if (data.Level == 2)
                {
                    var level3Data = (from wf in db.Tblgenericworkcodes
                                      where wf.IsDeleted == 0 && wf.GwcodesLevel2Id == data.GenericWorkId && wf.GwcodesLevel == 3
                                      select new
                                      {
                                          GenericWorkId = wf.GenericWorkId,
                                          GenericWorkCode = wf.GenericWorkCode,
                                          GenericWorkDesc = wf.GenericWorkDesc,
                                          GwcodesLevel = wf.GwcodesLevel,
                                          GwcodesLevel1Id = wf.GwcodesLevel1Id,
                                          GwcodesLevel2Id = wf.GwcodesLevel2Id
                                      }).ToList();
                    if (level3Data.Count == 0)
                    {
                        obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                        obj.BreakdownId = data.GenericWorkId;
                        obj.Level = data.Level;
                        obj.MachineDisplayName = machinedispname;
                    }
                    else
                    {
                        obj.Level = data.Level + 1;
                        obj.BreakdownId = data.GenericWorkId;
                        obj.MachineDisplayName = machinedispname;
                        obj.isStatus = true;
                        obj.response = level3Data;
                    }
                }
                else if (data.Level == 3)
                {
                    int prevLevelId = Convert.ToInt32(lossdata.GwcodesLevel2Id);
                    var level2Data = (from wf in db.Tblgenericworkcodes
                                      where wf.IsDeleted == 0 && wf.GwcodesLevel2Id == prevLevelId && wf.GwcodesLevel == 3
                                      select new
                                      {
                                          GenericWorkId = wf.GenericWorkId,
                                          GenericWorkCode = wf.GenericWorkCode,
                                          GenericWorkDesc = wf.GenericWorkDesc,
                                          GwcodesLevel = wf.GwcodesLevel,
                                          GwcodesLevel1Id = wf.GwcodesLevel1Id,
                                          GwcodesLevel2Id = wf.GwcodesLevel2Id
                                      }).ToList();
                    obj.ItsLastLevel = "No Further Levels . Do you want to set " + losscode + " as reason.";
                    obj.BreakdownId = data.GenericWorkId;
                    obj.Level = 3;
                    obj.isStatus = true;
                    obj.MachineDisplayName = machinedispname;
                    obj.response = level2Data;
                }
            }
            else
            {
                var level1Data = (from wf in db.Tblgenericworkcodes
                                  where wf.IsDeleted == 0 && wf.GwcodesLevel == 1
                                  select new
                                  {
                                      GenericWorkId = wf.GenericWorkId,
                                      GenericWorkCode = wf.GenericWorkCode,
                                      GenericWorkDesc = wf.GenericWorkDesc,
                                      GwcodesLevel = wf.GwcodesLevel,
                                      GwcodesLevel1Id = wf.GwcodesLevel1Id,
                                      GwcodesLevel2Id = wf.GwcodesLevel2Id
                                  }).ToList();
                if (level1Data.Count != 0)
                {
                    obj.isStatus = true;
                    obj.response = level1Data;
                    obj.Level = 1;
                    obj.MachineDisplayName = machinedispname;
                }
            }
            return obj;
        }

        public CommonResponse UploadDDL(DdlDetails data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                foreach (var item in data.ddls)
                {
                    var check = db.TblSpddl.Where(m => m.WorkOrder == item.ProductionOrder && m.OperationNo == item.OperationNo && m.PartName == item.PartNumber).FirstOrDefault();

                    #region
                    string CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                    if (DateTime.Now.Hour < 6 && DateTime.Now.Hour >= 0)
                    {
                        CorrectedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                    }
                    #endregion

                    if (check == null)
                    {
                        TblSpddl tblSpddl = new TblSpddl();
                        tblSpddl.WorkCenter = item.WorkCenter;
                        tblSpddl.WorkOrder = item.ProductionOrder;
                        tblSpddl.PartName = item.PartNumber;
                        tblSpddl.OperationNo = item.OperationNo;
                        tblSpddl.OperationDesc = item.Operationshortdesc;
                        tblSpddl.MaterialDesc = item.MatrialDescription;
                        tblSpddl.TargetQty = item.TargetQty;
                        tblSpddl.Type = item.OrderCatgory;
                        tblSpddl.Project = item.Project;
                        tblSpddl.Maddate = item.MADDate;
                        tblSpddl.MaddateInd = item.MADDateInd;
                        tblSpddl.PreOpnEndDate = item.PreOpnEndDate;
                        tblSpddl.DaysAgeing = item.DaysAgeing;
                        tblSpddl.OperationsOnHold = item.OperationsOnHold;
                        tblSpddl.ReasonForHold = item.ReasonForHold;
                        tblSpddl.DueDate = item.DueDate;
                        tblSpddl.FlagRushInd = item.FiagRushInd;
                        tblSpddl.SplitWo = item.SpiltWorkOrder;
                        tblSpddl.IsCompleted = 0;
                        tblSpddl.IsDeleted = 0;
                        tblSpddl.IsWoselected = 0;
                        tblSpddl.InsertedOn = DateTime.Now;
                        tblSpddl.CorrectedDate = CorrectedDate;
                        tblSpddl.Bhmiid = 0;
                        tblSpddl.DeliveredQty = 0;
                        tblSpddl.ScrapQty = 0;
                        db.TblSpddl.Add(tblSpddl);
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Added Successully";
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = "Data Already Present";
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Newly Uploaded Ddl Details
        /// </summary>
        /// <returns></returns>
        public CommonResponseCountList GetNewlyUploadedDdlDetails(int MachineId)
        {
            CommonResponseCountList obj = new CommonResponseCountList();
            var machinedispname = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == MachineId).Select(m => m.MachineDispName).FirstOrDefault();
            try
            {
                var check = (from wf in db.TblSpddl
                             where wf.IsDeleted == 0 && wf.IsCompleted == 0 && wf.IsWoselected == 0
                             select new
                             {
                                 SpDdlid = wf.SpDdlid,
                                 WorkCenter = wf.WorkCenter,
                                 WorkOrder = wf.WorkOrder,
                                 Type = wf.Type,
                                 TargetQty = wf.TargetQty,
                                 SplitWo = wf.SplitWo,
                                 ScrapQty = wf.ScrapQty,
                                 ReasonForHold = wf.ReasonForHold,
                                 Project = wf.Project,
                                 PreOpnEndDate = wf.PreOpnEndDate,
                                 PartName = wf.PartName,
                                 OperationsOnHold = wf.OperationsOnHold,
                                 OperationNo = wf.OperationNo,
                                 OperationDesc = wf.OperationDesc,
                                 MaterialDesc = wf.MaterialDesc,
                                 MaddateInd = wf.MaddateInd,
                                 Maddate = wf.Maddate,
                                 FlagRushInd = wf.FlagRushInd,
                                 DeliveredQty = wf.DeliveredQty,
                                 DaysAgeing = wf.DaysAgeing,
                                 CorrectedDate = wf.CorrectedDate,
                                 DueDate = wf.DueDate
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                    obj.MachineDisplayName = machinedispname;
                    obj.Count = check.Count();
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

        public CommonResponsewithEror OperatorLoginCheck(Oplogin data)
        {
            CommonResponsewithEror obj = new CommonResponsewithEror();
            try
            {
                var modeData = db.Tblusers.Where(m => m.UserId == data.operatorId && m.Password == data.password && m.IsDeleted == 0).FirstOrDefault();
                if (modeData != null)
                {
                    obj.isTure = true;
                    obj.response = "Opertor Login Correct";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No Data Found";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }


        


        

        #endregion
    }
}
#endregion
