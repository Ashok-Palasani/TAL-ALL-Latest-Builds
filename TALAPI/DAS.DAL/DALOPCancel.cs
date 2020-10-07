using DAS.DAL.Resource;
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
using static DAS.EntityModels.OPCancelEntity;
using static DAS.EntityModels.SplitDurationEntity;

namespace DAS.DAL
{
    public class DALOPCancel : IOpCancel
    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALLoss));

        public static IConfiguration configuration;

        public DALOPCancel(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }

        public CommonResponse GetPrevOpCancelDetails()
        {
            CommonResponse entity = new CommonResponse();

            try
            {
                List<TblPrevOperationCancel> PrevOpCancelDetails = new List<TblPrevOperationCancel>();

            }
            catch (Exception ex)
            {

            }
            return entity;
        }


        //public EntityModel GetPrevOpCancelDetails(string prevoperationDets)
        //{
        //    EntityModel entity = new EntityModel();

        //    try
        //    {
        //       // List<TblPrevOperationCancel> opcancelfromexcel= 
        //        List<TblPrevOperationCancel> PrevOpCancelDetails = new List<TblPrevOperationCancel>();

        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return entity;
        //}

        public CommonResponse GetPrevOpCancelDetails(LsitOPcancelDet Opcanceldet)
        {
            CommonResponse entity = new CommonResponse();

            try
            {
                if (Opcanceldet.OPCancelDetailsList.Count > 0)
                {
                    //string correctedDate = GetCorrectedDate();
                    //List<Tblhmiscreen> HMIList = new List<Tblhmiscreen>();

                    //HMIList = db.Tblhmiscreen.Where(m => m.CorrectedDate == correctedDate && m.Status==2).ToList();  // Jobfinished workorders

                    //var workordernum = HMIList.Select(m => m.WorkOrderNo).ToList(); 
                    //var operationnum = HMIList.Select(m => m.OperationNo).ToList();

                    //Opcanceldet.OPCancelDetailsList = Opcanceldet.OPCancelDetailsList.Where(m => m.CorrectedDate == correctedDate && workordernum.Contains(m.ProductionOrder) && operationnum.Contains(m.Operation)).ToList();


                }
                //List<TblPrevOperationCancel> PrevOpCancelDetails = new List<TblPrevOperationCancel>();

                entity.isTure = true;
                entity.response = Opcanceldet;
            }
            catch (Exception ex)
            {

            }
            return entity;
        }


        #region Vignesh logic for operation cancallation

        //For getting data from list and add into table
        public CommonResponse GetListOFOperationNumberFileData(List<UpLoadExcel> data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                SELOperationCancelDet objSELOperationCancelDet = new SELOperationCancelDet();
                List<OperationCancelDet> listOperationCancelDetSucess = new List<OperationCancelDet>();
                List<OperationCancelDet> listOperationCancelDetFail = new List<OperationCancelDet>();
                if (data.Count > 0)
                {
                    foreach (var row in data)
                    {
                        string workOrderNo = row.ProductionOrder;
                        string operationNo = row.Operation;
                        string partNo = row.PartNumber;
                        string workCenter = row.WorkCenter;
                        int machineId = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineInvNo == workCenter).Select(x => x.MachineId).FirstOrDefault();
                        var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == machineId && x.PartNo == partNo && x.OperationNo == operationNo && x.WorkOrderNo == workOrderNo && x.Status == 2 && x.IsWorkInProgress == 1).FirstOrDefault();
                        if (hmiData != null)
                        {
                            OperationCancelDet objOperationCancelDet = new OperationCancelDet();
                            objOperationCancelDet.correctedDate = row.CorrectedDate;
                            objOperationCancelDet.isCancelled = 0;
                            objOperationCancelDet.operation = operationNo;
                            objOperationCancelDet.partNo = partNo;
                            objOperationCancelDet.processedQty = row.Qty;
                            objOperationCancelDet.productionOrder = workOrderNo;
                            objOperationCancelDet.workCenter = workCenter;
                            listOperationCancelDetSucess.Add(objOperationCancelDet);

                        }
                        else
                        {
                            OperationCancelDet objOperationCancelDet = new OperationCancelDet();
                            objOperationCancelDet.correctedDate = row.CorrectedDate;
                            objOperationCancelDet.isCancelled = 0;
                            objOperationCancelDet.operation = operationNo;
                            objOperationCancelDet.partNo = partNo;
                            objOperationCancelDet.processedQty = row.Qty;
                            objOperationCancelDet.productionOrder = workOrderNo;
                            objOperationCancelDet.workCenter = workCenter;
                            listOperationCancelDetFail.Add(objOperationCancelDet);
                        }
                    }
                    InsertToTablePrvOperation(listOperationCancelDetSucess);// insert all data into table
                    List<OperationCancelDet> listOperationCancelDet = GetListOFOperationCancalaation();//get all the previous opretion for cancalation
                    objSELOperationCancelDet.OPCancelDetailsSuccessList = listOperationCancelDet;
                    objSELOperationCancelDet.OPCancelDetailsErrorList = listOperationCancelDetFail;

                    obj.isTure = true;
                    obj.response = objSELOperationCancelDet;

                }
                else
                {
                    obj.isTure = false;
                    obj.response = Resource.ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString());
                if (ex.InnerException.ToString() != null)
                {
                    log.Error(ex.InnerException.ToString());
                }
            }
            return obj;
        }

        // Insert into Tbale and get the list to show case
        public bool InsertToTablePrvOperation(List<OperationCancelDet> listOperationCancelDetSucess)
        {
            bool check = false;
            int updateLevel = 1;
            try
            {
                if (listOperationCancelDetSucess.Count > 0)
                {
                    foreach (var row in listOperationCancelDetSucess)
                    {
                        string workOrderNo = row.productionOrder;
                        string operationNo = row.operation;
                        string partNo = row.partNo;
                        string workCenter = row.workCenter;
                        //var prvOpCancelData = db.TblPrevOperationCancel.Where(x => x.PartNumber == partNo && x.Operation == operationNo && x.ProductionOrder == workOrderNo && x.WorkCenter == workCenter).FirstOrDefault();
                        //if (prvOpCancelData == null)
                        //{
                        var getMailIdsLevel = db.TblTcfApprovedMaster.Where(x => x.TcfModuleId == 6 && x.IsDeleted == 0).ToList();
                        foreach (var rowMail in getMailIdsLevel)
                        {
                            if (rowMail.SecondApproverCcList != null && rowMail.SecondApproverToList != null)
                            {
                                updateLevel = 2;
                            }
                        }
                        TblTcfPrevOperationCancel addRow = new TblTcfPrevOperationCancel();
                        addRow.CorrectedDate = row.correctedDate;
                        addRow.CreatedBy = 1;// change as per user login
                        addRow.IsCancelled = 0;
                        addRow.CreatedOn = DateTime.Now;
                        addRow.Operation = row.operation;
                        addRow.PartNumber = row.partNo;
                        addRow.ProductionOrder = row.productionOrder;
                        addRow.Qty = row.processedQty;
                        addRow.WorkCenter = row.workCenter;
                        addRow.UploadDate = DateTime.Now.ToString("yyyy-MM-dd");
                        addRow.UpdateLevel = updateLevel;
                        db.TblTcfPrevOperationCancel.Add(addRow);
                        db.SaveChanges();
                        //}
                        //else
                        //{
                        //    prvOpCancelData.IsCancelled = 0;
                        //    prvOpCancelData.Qty = row.processedQty;
                        //    prvOpCancelData.CorrectedDate = row.correctedDate;
                        //    db.SaveChanges();
                        //}
                    }

                }
            }
            catch (Exception ex)
            {
                check = false;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return check;
        }

        //getting list of all data for cancelation
        public List<OperationCancelDet> GetListOFOperationCancalaation()
        {
            List<OperationCancelDet> listOperationCancelDet = new List<OperationCancelDet>();
            try
            {
                var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 0).ToList();
                foreach (var row in operationCancelData)
                {
                    int qty = 0;
                    if (row.Qty > 0)
                    {
                        qty = Convert.ToInt32(row.Qty);
                    }
                    OperationCancelDet objOperationCancelDet = new OperationCancelDet();
                    objOperationCancelDet.correctedDate = row.CorrectedDate;
                    objOperationCancelDet.isCancelled = row.IsCancelled;
                    objOperationCancelDet.operation = row.Operation;
                    objOperationCancelDet.partNo = row.PartNumber;
                    objOperationCancelDet.processedQty = qty;
                    objOperationCancelDet.productionOrder = row.ProductionOrder;
                    objOperationCancelDet.workCenter = row.WorkCenter;
                    objOperationCancelDet.opCancelID = row.TcfopcancelId;
                    listOperationCancelDet.Add(objOperationCancelDet);
                }
            }
            catch (Exception ex)
            {
                listOperationCancelDet = null;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return listOperationCancelDet;
        }

        //select and unselect the operatio number
        public CommonResponse AcceptRejectOperationNo(AcceptCancel data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int opCancelID = data.opCancelID;
                var updateRow = db.TblTcfPrevOperationCancel.Where(x => x.TcfopcancelId == opCancelID).FirstOrDefault();
                if (updateRow != null)
                {
                    if (data.isChecked == 1)
                    {
                        updateRow.IsCancelled = 1;
                    }
                    else if (data.isChecked == 0)
                    {
                        updateRow.IsCancelled = 0;
                    }
                    db.SaveChanges();
                    obj.isTure = true;
                    obj.response = Resource.ResourceResponse.SuccessMessage;
                }
                else
                {
                    obj.isTure = false;
                    obj.response = Resource.ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.ToString()); }
            }
            return obj;
        }

        //get the the uncanclled list not sent for approval
        public CommonResponse GetCancledList()
        {
            CommonResponse obj = new CommonResponse();
            List<OperationCancelDet> listOperationCancelDet = new List<OperationCancelDet>();
            try
            {
                var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 0 && x.SendApprove == 0).ToList();
                foreach (var row in operationCancelData)
                {
                    int qty = 0;
                    if (row.Qty > 0)
                    {
                        qty = Convert.ToInt32(row.Qty);
                    }
                    OperationCancelDet objOperationCancelDet = new OperationCancelDet();
                    objOperationCancelDet.correctedDate = row.CorrectedDate;
                    objOperationCancelDet.isCancelled = row.IsCancelled;
                    objOperationCancelDet.operation = row.Operation;
                    objOperationCancelDet.partNo = row.PartNumber;
                    objOperationCancelDet.processedQty = qty;
                    objOperationCancelDet.productionOrder = row.ProductionOrder;
                    objOperationCancelDet.workCenter = row.WorkCenter;
                    objOperationCancelDet.opCancelID = row.TcfopcancelId;
                    listOperationCancelDet.Add(objOperationCancelDet);
                }
                obj.isTure = true;
                obj.response = listOperationCancelDet;
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //get the the canclled list fixed
        public CommonResponse GetUnCancledList()
        {
            CommonResponse obj = new CommonResponse();
            List<OperationCancelDet> listOperationCancelDet = new List<OperationCancelDet>();
            try
            {
                string correctedDate = GetCorrectedDate();
                var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 0 && x.UploadDate == correctedDate).ToList();
                foreach (var row in operationCancelData)
                {
                    int qty = 0;
                    if (row.Qty > 0)
                    {
                        qty = Convert.ToInt32(row.Qty);
                    }
                    OperationCancelDet objOperationCancelDet = new OperationCancelDet();
                    objOperationCancelDet.correctedDate = row.CorrectedDate;
                    objOperationCancelDet.isCancelled = row.IsCancelled;
                    objOperationCancelDet.operation = row.Operation;
                    objOperationCancelDet.partNo = row.PartNumber;
                    objOperationCancelDet.processedQty = qty;
                    objOperationCancelDet.productionOrder = row.ProductionOrder;
                    objOperationCancelDet.workCenter = row.WorkCenter;
                    objOperationCancelDet.opCancelID = row.TcfopcancelId;
                    listOperationCancelDet.Add(objOperationCancelDet);
                }
                obj.isTure = true;
                obj.response = listOperationCancelDet;
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //get current day date
        public string GetCorrectedDate()
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
            return CorrectedDate;
        }

        //Send Approvall Method for accept and reject
        public CommonResponse SendForApproval()
        {
            CommonResponse obj = new CommonResponse();   // get all the details then send by url or by html
            try
            {
                string machName = "";
                var reader = Path.Combine(@"C:\TataReport\TCFTemplate\PRVOPTemplate1.html");
                string htmlStr = File.ReadAllText(reader);
                String[] seperator = { "{{WOStart}}" };
                string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                htmlStr = htmlStr.Replace("{{WOStart}}", "");
                htmlStr = htmlStr.Replace("{{WOEnd}}", "");
                string uploadDate = "";
                int sl = 1;
                int MachineID = 0;
                var prvOperationData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 0).OrderBy(m => m.TcfopcancelId).ToList();

                foreach (var row in prvOperationData)
                {
                    row.SendApprove = 1;
                    row.IsPending = 1;
                    db.SaveChanges();

                    var wodet = db.Tblhmiscreen.Where(m => m.WorkOrderNo == row.ProductionOrder && m.OperationNo == row.Operation && m.PartNo == row.PartNumber).FirstOrDefault();
                    if(wodet !=null)
                    {
                        MachineID = wodet.MachineId;
                    }
                    uploadDate = row.UploadDate;

                    String slno = Convert.ToString(sl);
                    htmlStr = htmlStr.Replace("{{slno}}", slno);
                    htmlStr = htmlStr.Replace("{{MachineName}}", row.WorkCenter);
                    htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.ProductionOrder);
                    htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNumber);
                    htmlStr = htmlStr.Replace("{{OprationNo}}", row.Operation);
                    htmlStr = htmlStr.Replace("{{QTY}}", Convert.ToString(row.Qty));
                    htmlStr = htmlStr.Replace("{{CorrectedDate}}", row.CorrectedDate);

                    if (prvOperationData.Count == 1)
                    {
                        htmlStr = htmlStr.Replace("{{WO}}", "");
                    }
                    else if (sl < prvOperationData.Count)
                    {

                        htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                    }
                    else
                    {
                        htmlStr = htmlStr.Replace("{{WO}}", woHtml);
                    }
                    sl++;
                }
                machName = "Operation Cancellation " + uploadDate;

                htmlStr = htmlStr.Replace(woHtml, "");
                htmlStr = htmlStr.Replace("{{secondLevel}}", "For 1st Level Approval");

                //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLPRVOP").Value;
                string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLPRVOP").Value;

                string rejectSrc = rejectUrl + "uploadDate=" + uploadDate + "&checked=0&MachineID=" + MachineID + "";
                //string acceptSrc = acceptUrl + "uploadDate=" + uploadDate;

                string toMailIds = "";
                string ccMailIds = "";

                var tcfApproveMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 6).ToList();
                foreach (var row in tcfApproveMail)
                {
                    toMailIds += row.FirstApproverToList + ",";
                    ccMailIds += row.FirstApproverCcList + ",";
                }

                htmlStr = htmlStr.Replace("{{WO}}", "");
                htmlStr = htmlStr.Replace("{{userName}}", "All");
                //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                toMailIds = toMailIds.Remove(toMailIds.Length - 1);// removing last comma
                ccMailIds = ccMailIds.Remove(ccMailIds.Length - 1);// removing last comma


                bool ret = SendMail(htmlStr, toMailIds, ccMailIds, 1, machName);

                if (ret)
                {
                    obj.isTure = true;
                    obj.response = "Sent Mail For Approval";
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

        //Accept tht previous operation cancallation even 2nd approval
        public CommonResponse AcceptPrvOp(ApprovalClass data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                bool updateReportTab = false;
                string machName = "";
                int appLevl = 0;
                string toMail = "", ccMail = "";
                string[] ids = data.id.Split(',');
                foreach (var idrow in ids)
                {
                    int previd = Convert.ToInt32(idrow);
                    var tcfPrvOpData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.Update == 0 && x.UploadDate == data.uploadDate && (x.AcceptReject == 0 || x.AcceptReject1 == 0) && x.TcfopcancelId == previd).FirstOrDefault();
                    if (tcfPrvOpData != null)
                    {
                        machName = "Operation Cancellation " + data.uploadDate;
                        //foreach (var row in tcfPrvOpData)
                        //{
                        if (tcfPrvOpData.AcceptReject == 0)
                        {
                            tcfPrvOpData.AcceptReject = 1;
                            tcfPrvOpData.IsPending = 0;
                            tcfPrvOpData.ApprovalLevel = 1;
                            db.SaveChanges();
                            if (tcfPrvOpData.UpdateLevel == 1) // checking the report table updation need or not
                            {
                                updateReportTab = true;
                                appLevl = 1;
                            }
                            else
                            {
                                appLevl = 2;
                            }
                        }
                        else if (tcfPrvOpData.AcceptReject1 == 0)
                        {
                            tcfPrvOpData.AcceptReject1 = 1;
                            tcfPrvOpData.IsPending = 0;
                            tcfPrvOpData.ApprovalLevel = 2;
                            db.SaveChanges();
                            if (tcfPrvOpData.UpdateLevel == 2)  // checking the report table updation need or not
                            {
                                updateReportTab = true;
                            }
                            //appLevl = 2;
                        }
                    }

                }
                if (data.unCheckId != "")
                {
                    string[] unCheckedids = data.unCheckId.Split(',');
                    foreach (var uncheckedIdRow in unCheckedids)
                    {
                        int id = Convert.ToInt32(uncheckedIdRow);
                        var getNoCodeDet = db.TblTcfPrevOperationCancel.Where(m => m.TcfopcancelId == id).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            getNoCodeDet.IsPending = 1;
                            db.SaveChanges();
                        }
                    }
                }

                var checkApprovalMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 6).ToList();
                foreach (var mailRow in checkApprovalMail)
                {
                    if (appLevl == 1)
                    {
                        //if (mailRow.FirstApproverToList != null && mailRow.FirstApproverCcList != null)
                        //{
                        toMail = mailRow.FirstApproverToList + ",";
                        ccMail = mailRow.FirstApproverCcList + ",";
                        // }
                    }
                    else if (appLevl == 2)
                    {
                        //if (mailRow.SecondApproverToList != null && mailRow.SecondApproverCcList != null)
                        //{
                        toMail = mailRow.SecondApproverToList + ",";
                        ccMail = mailRow.SecondApproverCcList + ",";
                        //}
                    }
                    else if (appLevl == 0)
                    {
                        toMail += mailRow.FirstApproverToList + ",";
                        ccMail += mailRow.FirstApproverCcList + ",";
                        if (mailRow.SecondApproverToList != "" || mailRow.SecondApproverToList != null)
                        {
                            toMail += mailRow.SecondApproverToList + ",";
                            ccMail += mailRow.SecondApproverCcList + ",";
                        }
                    }
                }


                if (updateReportTab)// send mail and update the report table
                {
                    //bool ret = false;
                    try
                    {
                        bool ret = false;
                        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\PRVOPTemplate1.html");
                        string htmlStr = File.ReadAllText(reader);
                        String[] seperator = { "{{WOStart}}" };
                        string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                        var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                        htmlStr = htmlStr.Replace("{{WOStart}}", "");
                        htmlStr = htmlStr.Replace("{{WOEnd}}", "");
                        int sl = 1;

                        var prvOperationData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.AcceptReject == 1 && x.AcceptReject1 == 1 && x.Update == 0 && x.UploadDate == data.uploadDate).OrderBy(m => m.TcfopcancelId).ToList();

                        foreach (var row in prvOperationData)
                        {
                            //row.SendApprove = 1;
                            //db.SaveChanges();

                            data.uploadDate = row.UploadDate;

                            String slno = Convert.ToString(sl);
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", row.WorkCenter);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.ProductionOrder);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNumber);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.Operation);
                            htmlStr = htmlStr.Replace("{{QTY}}", Convert.ToString(row.Qty));
                            htmlStr = htmlStr.Replace("{{CorrectedDate}}", row.CorrectedDate);

                            if (prvOperationData.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < prvOperationData.Count)
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
                        ret = SendMail(htmlStr, toMail, ccMail, 2, machName);
                        if (ret) //reportupdate
                        {
                            bool check = InsertIntoPrvOpTab(data.uploadDate);
                            if (check)
                            {
                                obj.isTure = true;
                                obj.response = "Sent Mail";
                            }
                            else
                            {
                                obj.isTure = false;
                                obj.response = ResourceResponse.FailureMessage;
                            }
                        }


                        //     string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
                        //"</head><body><p>Dear " + " All" + ",</p></br><p><center> The Operation Cancellation Has Been Accepted</center></p></br><p>Thank you" +
                        //"</p></body></html>";

                        //     bool ret = SendMail(message, toMail, ccMail, 0, machName);

                        //     //if (levl2toMail != "" && levl2ccMail != "")
                        //     //{
                        //     //    levl2toMail = levl2toMail.Remove(levl2toMail.Length - 1);
                        //     //    levl2ccMail = levl2ccMail.Remove(levl2ccMail.Length - 1);
                        //     //    ret = SendMail(message, levl2toMail, levl2ccMail, 0,machName);
                        //     //}
                        //     //else if (levl1toMail != "" && levl1ccMail != "")
                        //     //{
                        //     //    levl1toMail = levl1toMail.Remove(levl1toMail.Length - 1);
                        //     //    levl1ccMail = levl1ccMail.Remove(levl1ccMail.Length - 1);
                        //     //    ret = SendMail(message, levl1toMail, levl1ccMail, 0,machName);
                        //     //}


                        //     if (ret) //reportupdate
                        //     {
                        //         bool check = InsertIntoPrvOpTab(data.uploadDate);
                        //         if (check)
                        //         {
                        //             obj.isTure = true;
                        //             obj.response = "Data Updated for Reports";
                        //         }
                        //     }

                    }
                    catch (Exception ex)
                    {
                        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                    }
                }
                else // send mail
                {
                    try
                    {
                        bool ret = false;
                        var reader = Path.Combine(@"C:\TataReport\TCFTemplate\PRVOPTemplate1.html");
                        string htmlStr = File.ReadAllText(reader);
                        String[] seperator = { "{{WOStart}}" };
                        string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                        var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                        htmlStr = htmlStr.Replace("{{WOStart}}", "");
                        htmlStr = htmlStr.Replace("{{WOEnd}}", "");
                        int sl = 1;
                        int MachineID = 0;
                        var prvOperationData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.AcceptReject == 1 && x.AcceptReject1 == 0 && x.Update == 0 && x.UploadDate == data.uploadDate).OrderBy(m => m.TcfopcancelId).ToList();

                        foreach (var row in prvOperationData)
                        {
                            //row.SendApprove = 1;
                            //db.SaveChanges();
                            var wodet = db.Tblhmiscreen.Where(m => m.WorkOrderNo == row.ProductionOrder && m.OperationNo == row.Operation && m.PartNo == row.PartNumber).FirstOrDefault();
                            if (wodet != null)
                            {
                                MachineID = wodet.MachineId;
                            }
                            data.uploadDate = row.UploadDate;

                            String slno = Convert.ToString(sl);
                            htmlStr = htmlStr.Replace("{{slno}}", slno);
                            htmlStr = htmlStr.Replace("{{MachineName}}", row.WorkCenter);
                            htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.ProductionOrder);
                            htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNumber);
                            htmlStr = htmlStr.Replace("{{OprationNo}}", row.Operation);
                            htmlStr = htmlStr.Replace("{{QTY}}", Convert.ToString(row.Qty));
                            htmlStr = htmlStr.Replace("{{CorrectedDate}}", row.CorrectedDate);

                            if (prvOperationData.Count == 1)
                            {
                                htmlStr = htmlStr.Replace("{{WO}}", "");
                            }
                            else if (sl < prvOperationData.Count)
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

                        //string acceptUrl = configuration.GetSection("MySettings").GetSection("AcceptURLPRVOP").Value;
                        string rejectUrl = configuration.GetSection("MySettings").GetSection("RejectURLPRVOP").Value;

                        string rejectSrc = rejectUrl + "uploadDate=" + data.uploadDate + "&checked=" + data.id + "&MachineID=" + MachineID + "";
                        //string acceptSrc = acceptUrl + "uploadDate=" + data.uploadDate;

                        htmlStr = htmlStr.Replace("{{WO}}", "");
                        htmlStr = htmlStr.Replace("{{userName}}", "All");
                        //htmlStr = htmlStr.Replace("{{urlA}}", acceptSrc);
                        htmlStr = htmlStr.Replace("{{urlAR}}", rejectSrc);

                        ret = SendMail(htmlStr, toMail, ccMail, 1, machName);
                        //if (levl2toMail != "" && levl2ccMail != "")
                        //{
                        //    levl2toMail = levl2toMail.Remove(levl2toMail.Length - 1);
                        //    levl2ccMail = levl2ccMail.Remove(levl2ccMail.Length - 1);
                        //    ret = SendMail(htmlStr, levl2toMail, levl2ccMail, 1, machName);
                        //}
                        //else if (levl1toMail != "" && levl1ccMail != "")
                        //{
                        //    levl1toMail = levl1toMail.Remove(levl1toMail.Length - 1);
                        //    levl1ccMail = levl1ccMail.Remove(levl1ccMail.Length - 1);
                        //    ret = SendMail(htmlStr, levl1toMail, levl1ccMail, 1, machName);
                        //}

                        if (ret)
                        {
                            obj.isTure = true;
                            obj.response = "Sent For Second Level Approval";
                        }
                        else
                        {
                            obj.isTure = false;
                            obj.response = ResourceResponse.FailureMessage;
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                    }
                }



                //else
                //{
                //    obj.isTure = false;
                //    obj.response = ResourceResponse.NoItemsFound;
                //}
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //update to the table previous cancelation
        public bool InsertIntoPrvOpTab(string uploadDate)
        {
            bool result = false;
            try
            {
                var tcfPrvOpData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.UploadDate == uploadDate && (x.AcceptReject == 1 || x.AcceptReject1 == 1) && x.Update == 0).ToList();
                foreach (var row in tcfPrvOpData)
                {
                    TblPrevOperationCancel addRow = new TblPrevOperationCancel();
                    addRow.CorrectedDate = row.CorrectedDate;
                    addRow.CreatedBy = 1;//change as per login
                    addRow.CreatedOn = DateTime.Now;
                    addRow.IsCancelled = row.IsCancelled;
                    addRow.Operation = row.Operation;
                    addRow.PartNumber = row.PartNumber;
                    addRow.ProductionOrder = row.ProductionOrder;
                    addRow.Qty = row.Qty;
                    addRow.WorkCenter = row.WorkCenter;
                    row.Update = 1;
                    db.TblPrevOperationCancel.Add(addRow);
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return result;
        }

        //mailing concept for send mail
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

                    var smtpConn = db.Smtpdetails.Where(x => x.IsDeleted == true && x.TcfModuleId == 6).FirstOrDefault();
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
                    if(image == 2)
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

        //Common method for displaying the data in accept and reject
        public CommonResponse CommonDtatForDisplay(ApprovalClass data)
        {
            CommonResponse obj = new CommonResponse();
            string[] ids = data.id.Split(',');
            List<IndexOperationCancelDet> listOperationCancelDet = new List<IndexOperationCancelDet>();
            try
            {
                if (data.id == "")
                {
                    //var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.UploadDate == data.uploadDate && (x.AcceptReject == 0 || x.AcceptReject == 1) && (x.AcceptReject1 == 1 || x.AcceptReject1 == 0) && x.Update == 0).ToList();
                    var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.UploadDate == data.uploadDate && x.IsPending==1 && x.Update == 0).OrderBy(m => m.TcfopcancelId).ToList();
                    if (operationCancelData.Count > 0)
                    {
                        foreach (var row in operationCancelData)
                        {
                            int qty = 0;
                            string acceptreject = "No";
                            if (row.Qty > 0)
                            {
                                qty = Convert.ToInt32(row.Qty);
                            }
                            if (row.AcceptReject == 1 || row.AcceptReject1 == 1 && row.RejectReason1 == 0)
                            {
                                acceptreject = "Yes";
                            }



                            IndexOperationCancelDet objOperationCancelDet = new IndexOperationCancelDet();
                            objOperationCancelDet.correctedDate = row.CorrectedDate;
                            objOperationCancelDet.isCancelled = row.IsCancelled;
                            objOperationCancelDet.operation = row.Operation;
                            objOperationCancelDet.partNo = row.PartNumber;
                            objOperationCancelDet.processedQty = qty;
                            objOperationCancelDet.productionOrder = row.ProductionOrder;
                            objOperationCancelDet.workCenter = row.WorkCenter;
                            objOperationCancelDet.opCancelID = row.TcfopcancelId;
                            objOperationCancelDet.sentApp = "Yes";
                            objOperationCancelDet.acceptReject = acceptreject;
                            listOperationCancelDet.Add(objOperationCancelDet);
                        }
                        obj.isTure = true;
                        obj.response = listOperationCancelDet;
                    }
                    else if(operationCancelData.Count == 0)
                    {
                        obj.isTure = false;
                        obj.response = "All The Work Orders are Accepted";
                    }
                }
                else
                {
                    bool ret = false;
                    foreach (var rowid in ids)
                    {
                        int id = Convert.ToInt32(rowid);
                        var nocdet = db.TblTcfPrevOperationCancel.Where(x => x.TcfopcancelId == id && x.IsPending == 0).FirstOrDefault();
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
                    if(ret == true)
                    {
                        foreach (var rowid in ids)
                        {
                            int id = Convert.ToInt32(rowid);
                            var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.TcfopcancelId == id && x.IsPending == 0 && (x.AcceptReject == 1 || x.AcceptReject1 == 1) && x.ApprovalLevel !=2).FirstOrDefault();
                            if (operationCancelData!=null)
                            {
                                //foreach (var row in operationCancelData)
                                //{
                                    int qty = 0;
                                    string acceptreject = "No";
                                    if (operationCancelData.Qty > 0)
                                    {
                                        qty = Convert.ToInt32(operationCancelData.Qty);
                                    }
                                    if (operationCancelData.AcceptReject == 1 || operationCancelData.AcceptReject1 == 1 && operationCancelData.RejectReason1 == 0)
                                    {
                                        acceptreject = "Yes";
                                    }



                                    IndexOperationCancelDet objOperationCancelDet = new IndexOperationCancelDet();
                                    objOperationCancelDet.correctedDate = operationCancelData.CorrectedDate;
                                    objOperationCancelDet.isCancelled = operationCancelData.IsCancelled;
                                    objOperationCancelDet.operation = operationCancelData.Operation;
                                    objOperationCancelDet.partNo = operationCancelData.PartNumber;
                                    objOperationCancelDet.processedQty = qty;
                                    objOperationCancelDet.productionOrder = operationCancelData.ProductionOrder;
                                    objOperationCancelDet.workCenter = operationCancelData.WorkCenter;
                                    objOperationCancelDet.opCancelID = operationCancelData.TcfopcancelId;
                                    objOperationCancelDet.sentApp = "Yes";
                                    objOperationCancelDet.acceptReject = acceptreject;
                                    listOperationCancelDet.Add(objOperationCancelDet);
                               // }
                                obj.isTure = true;
                                obj.response = listOperationCancelDet;
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
                            var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.TcfopcancelId == id && x.IsPending == 1).FirstOrDefault();
                            if (operationCancelData != null)
                            {
                                //foreach (var row in operationCancelData)
                                //{
                                int qty = 0;
                                string acceptreject = "No";
                                if (operationCancelData.Qty > 0)
                                {
                                    qty = Convert.ToInt32(operationCancelData.Qty);
                                }
                                if (operationCancelData.AcceptReject == 1 || operationCancelData.AcceptReject1 == 1 && operationCancelData.RejectReason1 == 0)
                                {
                                    acceptreject = "Yes";
                                }



                                IndexOperationCancelDet objOperationCancelDet = new IndexOperationCancelDet();
                                objOperationCancelDet.correctedDate = operationCancelData.CorrectedDate;
                                objOperationCancelDet.isCancelled = operationCancelData.IsCancelled;
                                objOperationCancelDet.operation = operationCancelData.Operation;
                                objOperationCancelDet.partNo = operationCancelData.PartNumber;
                                objOperationCancelDet.processedQty = qty;
                                objOperationCancelDet.productionOrder = operationCancelData.ProductionOrder;
                                objOperationCancelDet.workCenter = operationCancelData.WorkCenter;
                                objOperationCancelDet.opCancelID = operationCancelData.TcfopcancelId;
                                objOperationCancelDet.sentApp = "Yes";
                                objOperationCancelDet.acceptReject = acceptreject;
                                listOperationCancelDet.Add(objOperationCancelDet);
                                // }
                                obj.isTure = true;
                                obj.response = listOperationCancelDet;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
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
                            var check1 = db.TblTcfPrevOperationCancel.Where(m => intArry.Contains(m.TcfopcancelId) && m.SendApprove == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check1.Count > 0)
                            {
                                var toMail1 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 6).Select(m => m.FirstApproverToList).FirstOrDefault();

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

                            var check2 = db.TblTcfPrevOperationCancel.Where(m => intArry.Contains(m.TcfopcancelId) && m.ApprovalLevel == 1).Select(m => new { m.SendApprove, m.ApprovalLevel }).GroupBy(m => new { m.SendApprove, m.ApprovalLevel }).ToList();
                            if (check2.Count > 0)
                            {
                                var toMail2 = db.TblTcfApprovedMaster.Where(m => m.CellId == dbCheck.CellId && m.ShopId == dbCheck.ShopId && m.PlantId == dbCheck.PlantId && m.TcfModuleId == 6).Select(m => m.SecondApproverToList).FirstOrDefault();
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

        //Get the reasons
        public CommonResponse GetRejectReason()
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
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }

        //Reject the operation cancellation
        public CommonResponse RejectPrvOPCan(RejectClass data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                int checkMail = 0;
                string toMails = "", ccMails = "";
                string machName = "";
                string[] ids = data.id.Split(',');
                foreach (var idrow in ids)
                {
                    int previd = Convert.ToInt32(idrow);
                    var tcfPrvOpData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.UploadDate == data.uploadDate && (x.AcceptReject == 0 || x.AcceptReject1 == 0) && x.Update == 0 && x.TcfopcancelId == previd).FirstOrDefault();
                    if (tcfPrvOpData != null)
                    {
                        //foreach (var row in tcfPrvOpData)
                        //{
                        machName = "Operation Cancellation " + data.uploadDate;
                        if (tcfPrvOpData.AcceptReject == 0)
                        {
                            tcfPrvOpData.AcceptReject = 2;
                            tcfPrvOpData.IsPending = 0;
                            tcfPrvOpData.RejectReason = data.reasonId;
                            tcfPrvOpData.ApprovalLevel = 1;
                            db.SaveChanges();
                            checkMail = 1;
                        }
                        else if (tcfPrvOpData.AcceptReject1 == 0)
                        {
                            tcfPrvOpData.AcceptReject1 = 2;
                            tcfPrvOpData.IsPending = 0;
                            tcfPrvOpData.RejectReason1 = data.reasonId;
                            tcfPrvOpData.ApprovalLevel = 2;
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
                        var getNoCodeDet = db.TblTcfPrevOperationCancel.Where(m => m.TcfopcancelId == id).FirstOrDefault();
                        if (getNoCodeDet != null)
                        {
                            getNoCodeDet.IsPending = 1;
                            db.SaveChanges();
                        }
                    }
                }
                var checkApprovalMail = db.TblTcfApprovedMaster.Where(x => x.IsDeleted == 0 && x.TcfModuleId == 6).ToList();
                foreach (var mailRow in checkApprovalMail)
                {
                    if (checkMail == 1)
                    {
                        //if (mailRow.FirstApproverToList != null && mailRow.FirstApproverCcList != null)
                        //{
                        toMails = mailRow.FirstApproverToList + ",";
                        ccMails = mailRow.FirstApproverCcList + ",";
                        // }
                    }
                    if (checkMail == 2)
                    {
                        toMails += mailRow.FirstApproverToList + ",";
                        ccMails += mailRow.FirstApproverCcList + ",";
                        if (mailRow.SecondApproverToList != "" || mailRow.SecondApproverToList != null)
                        {
                            toMails += mailRow.SecondApproverToList + ",";
                            ccMails += mailRow.SecondApproverCcList + ",";
                        }
                    }
                }

                toMails = toMails.Remove(toMails.Length - 1);
                ccMails = ccMails.Remove(ccMails.Length - 1);

                bool ret = false;
                try
                {
                    string rejectName = db.Tblrejectreason.Where(x => x.IsDeleted == 0 && x.Rid == data.reasonId).Select(x => x.RejectNameDesc).FirstOrDefault();

                    //bool ret = false;
                    var reader = Path.Combine(@"C:\TataReport\TCFTemplate\PRVOPTemplate1.html");
                    string htmlStr = File.ReadAllText(reader);
                    String[] seperator = { "{{WOStart}}" };
                    string[] htmlArr = htmlStr.Split(seperator, 2, StringSplitOptions.RemoveEmptyEntries);

                    var woHtml = htmlArr[1].Split(new String[] { "{{WOEnd}}" }, 2, StringSplitOptions.RemoveEmptyEntries)[0];
                    htmlStr = htmlStr.Replace("{{WOStart}}", "");
                    htmlStr = htmlStr.Replace("{{WOEnd}}", "");
                    int sl = 1;

                    var prvOperationData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1 && x.AcceptReject == 2 && x.AcceptReject1 == 2 && x.Update == 0 && x.UploadDate == data.uploadDate).OrderBy(m => m.TcfopcancelId).ToList();

                    foreach (var row in prvOperationData)
                    {
                        //row.SendApprove = 1;
                        //db.SaveChanges();

                        data.uploadDate = row.UploadDate;

                        String slno = Convert.ToString(sl);
                        htmlStr = htmlStr.Replace("{{slno}}", slno);
                        htmlStr = htmlStr.Replace("{{MachineName}}", row.WorkCenter);
                        htmlStr = htmlStr.Replace("{{WorkOrderNo}}", row.ProductionOrder);
                        htmlStr = htmlStr.Replace("{{PartNo}}", row.PartNumber);
                        htmlStr = htmlStr.Replace("{{OprationNo}}", row.Operation);
                        htmlStr = htmlStr.Replace("{{QTY}}", Convert.ToString(row.Qty));
                        htmlStr = htmlStr.Replace("{{CorrectedDate}}", row.CorrectedDate);

                        if (prvOperationData.Count == 1)
                        {
                            htmlStr = htmlStr.Replace("{{WO}}", "");
                        }
                        else if (sl < prvOperationData.Count)
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
                    htmlStr = htmlStr.Replace("{{secondLevel}}", "Thease Work Order for Operation Cancellation Has Been Rejected for this " + rejectName + ".");
                    
               //     string message = "<!DOCTYPE html><html xmlns = 'http://www.w3.org/1999/xhtml' ><head><title></title><link rel='stylesheet' type='text/css' href='//fonts.googleapis.com/css?family=Open+Sans'/>" +
               //"</head><body><p>Dear " + " All" + ",</p></br><p><center> The Operation Cancellation Has Been Rejected for this " + rejectName + ".</center></p></br><p>Thank you" +
               //"</p></body></html>";


                    ret = SendMail(htmlStr, toMails, ccMails, 2, machName);


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
                    log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
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

        //Index Method

        #region Old Index
        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    List<IndexOperationCancelDet> listOperationCancelDet = new List<IndexOperationCancelDet>();
        //    try
        //    {
        //        var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1).ToList();
        //        foreach (var row in operationCancelData)
        //        {
        //            int qty = 0;
        //            string acceptreject = "No";
        //            if (row.Qty > 0)
        //            {
        //                qty = Convert.ToInt32(row.Qty);
        //            }
        //            if (row.AcceptReject == 1 || row.AcceptReject1 == 1 && row.RejectReason1 == 0)
        //            {
        //                acceptreject = "Yes";
        //            }

        //            IndexOperationCancelDet objOperationCancelDet = new IndexOperationCancelDet();
        //            objOperationCancelDet.correctedDate = row.CorrectedDate;
        //            objOperationCancelDet.isCancelled = row.IsCancelled;
        //            objOperationCancelDet.operation = row.Operation;
        //            objOperationCancelDet.partNo = row.PartNumber;
        //            objOperationCancelDet.processedQty = qty;
        //            objOperationCancelDet.productionOrder = row.ProductionOrder;
        //            objOperationCancelDet.workCenter = row.WorkCenter;
        //            objOperationCancelDet.opCancelID = row.TcfopcancelId;
        //            objOperationCancelDet.sentApp = "Yes";
        //            objOperationCancelDet.acceptReject = acceptreject;
        //            listOperationCancelDet.Add(objOperationCancelDet);
        //        }
        //        obj.isTure = true;
        //        obj.response = listOperationCancelDet;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = Resource.ResourceResponse.ExceptionMessage;
        //        log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}
        #endregion

        //public CommonResponse Index()
        //{
        //    CommonResponse obj = new CommonResponse();
        //    List<IndexOperationCancelDet1> listOperationCancelDet = new List<IndexOperationCancelDet1>();
        //    try
        //    {
        //        var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1).ToList();
        //        foreach (var row in operationCancelData)
        //        {
        //            int qty = 0;
        //            string firstApproval = "", secondApproval = "";
        //            //string acceptreject = "No";
        //            if (row.Qty > 0)
        //            {
        //                qty = Convert.ToInt32(row.Qty);
        //            }
        //            //if (row.AcceptReject == 1 || row.AcceptReject1 == 1 && row.RejectReason1 == 0)
        //            //{
        //            //    acceptreject = "Yes";
        //            //}

        //            if (row.AcceptReject1 == 1)
        //            {
        //                secondApproval = "Mail Sent and Second Level is Aprroved";

        //            }
        //            else if (row.AcceptReject1 == 2)
        //            {
        //                secondApproval = "Mail Sent and Second Level is Rejected";

        //            }
        //            if (row.AcceptReject == 1)
        //            {
        //                firstApproval = "Mail Sent and First Level is Aprroved";
        //            }
        //            else if (row.AcceptReject == 2)
        //            {
        //                firstApproval = "Mail Sent and First Level is Rejected";
        //            }

        //            IndexOperationCancelDet1 objOperationCancelDet = new IndexOperationCancelDet1();
        //            objOperationCancelDet.correctedDate = row.CorrectedDate;
        //            objOperationCancelDet.isCancelled = row.IsCancelled;
        //            objOperationCancelDet.operation = row.Operation;
        //            objOperationCancelDet.partNo = row.PartNumber;
        //            objOperationCancelDet.processedQty = qty;
        //            objOperationCancelDet.productionOrder = row.ProductionOrder;
        //            objOperationCancelDet.workCenter = row.WorkCenter;
        //            objOperationCancelDet.opCancelID = row.TcfopcancelId;
        //            objOperationCancelDet.firstApproval = firstApproval;
        //            objOperationCancelDet.secondApproval = secondApproval;
        //            listOperationCancelDet.Add(objOperationCancelDet);
        //        }
        //        obj.isTure = true;
        //        obj.response = listOperationCancelDet;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj.isTure = false;
        //        obj.response = Resource.ResourceResponse.ExceptionMessage;
        //        log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return obj;
        //}

        public CommonResponse Index()
        {
            CommonResponse obj = new CommonResponse();
            List<IndexOperationCancelDet1> listOperationCancelDet = new List<IndexOperationCancelDet1>();
            try
            {
                var operationCancelData = db.TblTcfPrevOperationCancel.Where(x => x.IsCancelled == 1 && x.SendApprove == 1).ToList();
                foreach (var row in operationCancelData)
                {
                    int qty = 0;
                    string firstApproval = "", secondApproval = "";
                    //string acceptreject = "No";
                    if (row.Qty > 0)
                    {
                        qty = Convert.ToInt32(row.Qty);
                    }
                    //if (row.AcceptReject == 1 || row.AcceptReject1 == 1 && row.RejectReason1 == 0)
                    //{
                    //    acceptreject = "Yes";
                    //}

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

                    IndexOperationCancelDet1 objOperationCancelDet = new IndexOperationCancelDet1();
                    objOperationCancelDet.correctedDate = row.CorrectedDate;
                    objOperationCancelDet.isCancelled = row.IsCancelled;
                    objOperationCancelDet.operation = row.Operation;
                    objOperationCancelDet.partNo = row.PartNumber;
                    objOperationCancelDet.processedQty = qty;
                    objOperationCancelDet.productionOrder = row.ProductionOrder;
                    objOperationCancelDet.workCenter = row.WorkCenter;
                    objOperationCancelDet.opCancelID = row.TcfopcancelId;
                    objOperationCancelDet.firstApproval = firstApproval;
                    objOperationCancelDet.secondApproval = secondApproval;
                    listOperationCancelDet.Add(objOperationCancelDet);
                }
                obj.isTure = true;
                obj.response = listOperationCancelDet;
            }
            catch (Exception ex)
            {
                obj.isTure = false;
                obj.response = Resource.ResourceResponse.ExceptionMessage;
                log.Error(ex.ToString()); if (ex.InnerException.ToString() != null) { log.Error(ex.InnerException.ToString()); }
            }
            return obj;
        }


        #endregion
    }
}
