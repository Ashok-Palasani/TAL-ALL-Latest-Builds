using DAS.DAL.Helpers;
using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace DAS.DAL
{
    public class DALReport : IReport
    {
        i_facility_talContext db = new i_facility_talContext();
        public static IConfiguration configuration;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALReport));

        public DALReport(i_facility_talContext _db, IConfiguration _configuration)//, AppSettings _appSettings
        {
            db = _db;
            configuration = _configuration;
            //appSettings = _appSettings;
        }

        //MachineStatusRegister Generation
        public CommonResponse MachineStatusRegister(EntityReport data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                var getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.IsNormalWc == 0).ToList();
                if (data.machineId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId && x.IsNormalWc == 0).ToList();
                }
                else if (data.cellId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId && x.IsNormalWc == 0).ToList();
                }
                else if (data.shopId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId && x.IsNormalWc == 0).ToList();
                }
                else if (data.plantId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId && x.IsNormalWc == 0).ToList();
                }
                DateTime startDate = Convert.ToDateTime(data.fromDate);
                DateTime endDate = Convert.ToDateTime(data.toDate);
                double dateDifference = endDate.Subtract(startDate).TotalDays;


                FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\MachineStatusReport.xlsx");

                ExcelPackage templatep = new ExcelPackage(templateFile);
                ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];

                String FileDir = @"C:\TataReport\ReportsList\" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd");
                bool exists = System.IO.Directory.Exists(FileDir);
                if (!exists)
                    System.IO.Directory.CreateDirectory(FileDir);

                FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
                if (newFile.Exists)
                {
                    try
                    {
                        newFile.Delete();  // ensures we create a new workbook
                        newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx"));
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.ToString());
                    }
                }
                //Using the File for generation and populating it
                ExcelPackage p = null;
                p = new ExcelPackage(newFile);
                ExcelWorksheet worksheet = null;

                //Creating the WorkSheet for populating
                try
                    {
                        worksheet = p.Workbook.Worksheets.Add(Convert.ToDateTime(data.fromDate).ToString("dd-MM-yyyy"), Templatews);
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.ToString());
                    }

                if (worksheet == null)
                {
                    worksheet = p.Workbook.Worksheets.Add(Convert.ToDateTime(data.fromDate).ToString("dd-MM-yyyy") + "1", Templatews);
                }
                int sheetcount = p.Workbook.Worksheets.Count;
                p.Workbook.Worksheets.MoveToStart(sheetcount-1);
                worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                worksheet.Cells["C3"].Value = data.fromDate;
                worksheet.Cells["E3"].Value = data.toDate;

                int startRow = 5;
                int slNo = 1;

                for (int i = 0; i <= dateDifference; i++)
                {
                    DateTime cDate = startDate.AddDays(i);
                    string correctedDate = cDate.ToString("yyyy-MM-dd");
                    foreach (var macRow in getMachineDetails)
                    {
                        string plantName = db.Tblplant.Where(x => x.IsDeleted == 0 && x.PlantId == macRow.PlantId).Select(x => x.PlantName).FirstOrDefault();
                        string shopName = db.Tblshop.Where(x => x.IsDeleted == 0 && x.ShopId == macRow.ShopId).Select(x => x.ShopName).FirstOrDefault();
                        string cellName = db.Tblcell.Where(x => x.IsDeleted == 0 && x.CellId == macRow.CellId).Select(x => x.CellName).FirstOrDefault();
                        string machineInvNo = macRow.MachineInvNo;
                        string MachineDisplayName = macRow.MachineDispName;
                        int machineId = macRow.MachineId;
                        var modeDetails = db.Tblmode.Where(x => x.IsDeleted == 0 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.IsCompleted == 1).ToList();
                        foreach (var modeRow in modeDetails)
                        {
                            DateTime modeStartTime = Convert.ToDateTime(modeRow.StartTime);
                            DateTime modeEndTime = Convert.ToDateTime(modeRow.EndTime);
                            //string modeDuration = Convert.ToDateTime(modeRow.DurationInSec).ToLocalTime().ToString("HH:mm:ss");
                            int duration = (int)modeRow.DurationInSec;
                            TimeSpan t = TimeSpan.FromSeconds(duration);                                                       

                            string shift = GetShift(modeEndTime);

                            string modeDuration = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                            t.Hours,
                                            t.Minutes,
                                            t.Seconds);


                            string mode = modeRow.Mode;
                            if (modeRow.Mode == "IDLE" && modeRow.DurationInSec > 120)
                            {
                                var lossOfEntryDetails = db.Tbllossofentry.Where(x => x.DoneWithRow == 1 && x.MachineId == machineId && x.CorrectedDate == correctedDate && x.StartDateTime >= modeRow.StartTime && x.EndDateTime <= modeRow.EndTime).ToList();
                                foreach (var lossrow in lossOfEntryDetails)
                                {
                                    DateTime lossStartTime = Convert.ToDateTime(lossrow.StartDateTime);
                                    DateTime lossEndTime = Convert.ToDateTime(lossrow.EndDateTime);
                                    int lDuration = (int)lossEndTime.Subtract(lossStartTime).TotalSeconds;
                                    TimeSpan t1 = TimeSpan.FromSeconds(lDuration);

                                    string lossDuration = string.Format("{0:D2}:{1:D2}:{2:D2}",
                                                    t1.Hours,
                                                    t1.Minutes,
                                                    t1.Seconds);
                                    int lossID = lossrow.MessageCodeId;
                                    if (lossrow != null)
                                    {
                                        worksheet.Cells["A" + startRow].Value = slNo;
                                        worksheet.Cells["B" + startRow].Value = plantName;
                                        worksheet.Cells["C" + startRow].Value = shopName;
                                        worksheet.Cells["D" + startRow].Value = cellName;
                                        worksheet.Cells["E" + startRow].Value = machineInvNo;
                                        worksheet.Cells["F" + startRow].Value = MachineDisplayName;
                                        worksheet.Cells["G" + startRow].Value = correctedDate;
                                        worksheet.Cells["H" + startRow].Value = shift;
                                        worksheet.Cells["I" + startRow].Value = lossStartTime;
                                        worksheet.Cells["J" + startRow].Value = lossEndTime;
                                        worksheet.Cells["K" + startRow].Value = lossDuration;
                                        worksheet.Cells["L" + startRow].Value = modeRow.Mode;


                                        string[] lossDesc = LossLevel(lossID).Split(',');

                                        int lossPrint = 13;
                                        for (int j = 0; j < lossDesc.Count(); j++)
                                        {
                                            string columnName = ExcelColumnFromNumber(lossPrint);
                                            worksheet.Cells[columnName + startRow].Value = lossDesc[j].ToString();
                                            lossPrint++;
                                        }

                                    }
                                    slNo++;
                                    startRow++;
                                }
                            }
                            else
                            {
                                if (mode == "PowerOn")
                                {
                                    mode = "Production";
                                }
                                worksheet.Cells["A" + startRow].Value = slNo;
                                worksheet.Cells["B" + startRow].Value = plantName;
                                worksheet.Cells["C" + startRow].Value = shopName;
                                worksheet.Cells["D" + startRow].Value = cellName;
                                worksheet.Cells["E" + startRow].Value = machineInvNo;
                                worksheet.Cells["F" + startRow].Value = MachineDisplayName;
                                worksheet.Cells["G" + startRow].Value = correctedDate;
                                worksheet.Cells["H" + startRow].Value = shift;
                                worksheet.Cells["I" + startRow].Value = modeStartTime;
                                worksheet.Cells["J" + startRow].Value = modeEndTime;
                                worksheet.Cells["K" + startRow].Value = modeDuration;
                                worksheet.Cells["L" + startRow].Value = mode;
                                worksheet.Cells["M" + startRow].Value = "-";
                                worksheet.Cells["N" + startRow].Value = "-";
                                worksheet.Cells["O" + startRow].Value = "-";
                                slNo++;
                                startRow++;
                            }
                        }
                    }
                }



                //p.Workbook.Worksheets.MoveToStart(1);
                p.Save();

                //Downloding Excel
                string path1 = System.IO.Path.Combine(FileDir, "MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx");
                string destinationPath = @"C:\TCFReport\MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx";

                //getting the report ip address from app string.json
                string reportURL = configuration.GetSection("MySettings").GetSection("ReportURL").Value;

                File.Copy(path1, destinationPath, true);
                path1 = reportURL + "MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx";
                //path1 = @"http://192.168.0.16:8083/MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx";
                //path1 = @"http://192.168.0.7:8083/MachineStatusRegister" + Convert.ToDateTime(data.fromDate).ToString("yyyy-MM-dd") + ".xlsx";
                obj.isTure = true;
                obj.response = path1;
            }
            catch (Exception ex)
            {
                log.Error(ex.InnerException.ToString());
                obj.isTure = false;
                obj.response = ResourceResponse.ExceptionMessage;
            }

            return obj;
        }

        //ManualWOConfirmationAndSplit Generation
        //public CommonResponse ManualWOConfirmationAndSplitAndStart(ReportEntity data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        if (data.Type == "Start")
        //        {
        //            string finalPathWithFile = "";
        //            try
        //            {
        //                string Todate = DateTime.Now.ToString("yyyy-MM-dd");
        //                var OutputPath = configuration.GetSection("MySettings").GetSection("ImageUrlSave").Value;

        //                string fileNameIn = OutputPath + "WOStart" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
        //                //var OutputPath = db.Tblpreactorschedule.Where(x => x.IsDeleted == 0 && x.IsStart == 1).FirstOrDefault();

        //                //string fileNameIn = OutputPath.Path + OutputPath.FileFormat + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";

        //                string folderName = OutputPath + "WOStart/";

        //                if (!Directory.Exists(folderName))
        //                {
        //                    Directory.CreateDirectory(folderName);
        //                }

        //                // Check if file already exists. If yes, delete it.     
        //                if (File.Exists(fileNameIn))
        //                {
        //                    File.Delete(fileNameIn);
        //                }

        //                int startRow = 2;
        //                string fileName = "";
        //                string writeData = "";

        //                DateTime startDate = DateTime.Now;


        //                var wodet = db.Tbllivehmiscreen.Where(m => m.Date <= startDate && m.Time == null && m.IsWorkInProgress == 2 && m.CorrectedDate == Todate).ToList();
        //                foreach (var hmirow in wodet)
        //                {
        //                    var macdet = db.Tblmachinedetails.Where(m => m.MachineId == hmirow.MachineId).FirstOrDefault();

        //                    fileName = "WoStart" + "_" + "ZCO11" + "_" + startDate.ToString("yyyy-MM-dd").Replace("-", "") + startDate.ToString("HH:mm:ss").Replace(":", "");

        //                    var plantCode = db.Tblplant.Where(m => m.PlantId == macdet.PlantId).Select(m => m.PlantCode).FirstOrDefault();
        //                    int pendingqty = Convert.ToInt32(hmirow.TargetQty - hmirow.DeliveredQty - hmirow.RejQty);
        //                    DateTime StartDate = (DateTime)hmirow.Date;
        //                    string stDate = StartDate.ToString("yyyy-MM-dd");
        //                    string stTime = StartDate.ToString("HH:mm:ss");

        //                    string OperatorID = hmirow.OperatorDet;
        //                    if (OperatorID.Contains(","))
        //                    {
        //                        string[] data1 = OperatorID.Split(',');
        //                        OperatorID = data1[0];
        //                    }

        //                    startRow++;

        //                    fileName = fileName.Replace("-", "");
        //                    fileName = fileName.Replace(":", "");


        //                    finalPathWithFile = System.IO.Path.Combine(folderName, fileName + ".txt");



        //                    // Create a new file     

        //                    using (StreamWriter writer = new StreamWriter(@finalPathWithFile, true)) //true => Append Text
        //                    {
        //                        writeData = macdet.MachineInvNo + "\t" + plantCode + "\t" + OperatorID
        //                              + "\t" + hmirow.WorkOrderNo + "\t" + hmirow.OperationNo + "\t" + pendingqty + "\t" + stDate + "\t" +
        //                              stTime;
        //                        writer.WriteLine(writeData);
        //                    }

        //                }

        //                //getting the report ip address from app string.json
        //                string reportURL = configuration.GetSection("MySettings").GetSection("ImageUrl").Value;
        //                string path1 = reportURL + "WOStart/" + fileName + ".txt";

        //                if (fileName != "")
        //                {
        //                    obj.isTure = true;
        //                    obj.response = path1;

        //                }
        //                else
        //                {
        //                    obj.isTure = false;
        //                    obj.response = "No Data Found";
        //                }

        //            }
        //            catch (Exception ex)
        //            {
        //                log.Error(ex.InnerException.ToString());
        //                obj.isTure = false;
        //                obj.response = ResourceResponse.ExceptionMessage;
        //            }
        //        }
        //        else if (data.Type == "Confirmation")
        //        {
        //            string finalPathWithFile = "";
        //            try
        //            {
        //                DateTime startDate = Convert.ToDateTime(data.Fromdate);
        //                DateTime endDate = Convert.ToDateTime(data.ToDate);
        //                double dateDifference = endDate.Subtract(startDate).TotalDays;

        //                var OutputPath = configuration.GetSection("MySettings").GetSection("ImageUrlSave").Value;

        //                string fileNameIn = OutputPath + "WOConfirmation" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
        //                //var OutputPath = db.Tblpreactorschedule.Where(x => x.IsDeleted == 0 && x.IsStart == 1).FirstOrDefault();

        //                //string fileNameIn = OutputPath.Path + OutputPath.FileFormat + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";

        //                string folderName = OutputPath + "WOConfirmation/";

        //                if (!Directory.Exists(folderName))
        //                {
        //                    Directory.CreateDirectory(folderName);
        //                }
        //                // Check if file already exists. If yes, delete it.     
        //                if (File.Exists(fileNameIn))
        //                {
        //                    File.Delete(fileNameIn);
        //                }

        //                int startRow = 2;
        //                string fileName = "";
        //                string writeData = "";


        //                DataTable dtNew = new DataTable();
        //                dtNew.Columns.Add("WONo", typeof(string));
        //                dtNew.Columns.Add("OpNo", typeof(int));
        //                dtNew.Columns.Add("Qty", typeof(int));
        //                dtNew.Columns.Add("Scrap", typeof(int));
        //                dtNew.Columns.Add("SettingTime", typeof(int));
        //                dtNew.Columns.Add("MachiningTime", typeof(int));
        //                dtNew.Columns.Add("LabourTime", typeof(int));
        //                dtNew.Columns.Add("StartDate", typeof(string));
        //                dtNew.Columns.Add("StartTime", typeof(string));
        //                dtNew.Columns.Add("EndDate", typeof(string));
        //                dtNew.Columns.Add("EndTime", typeof(string));
        //                dtNew.Columns.Add("OperatorID", typeof(string));

        //                DateTime StartTimeNew = startDate;
        //                DateTime EndTimeNew = endDate;

        //                var GetJFSingleWO = db.Tbllivehmiscreen.Where(m => m.Time > StartTimeNew && m.Time <= EndTimeNew && m.IsWorkInProgress == 1).ToList();

        //                //Code for Single JO for Standard Work Centers Job Finish
        //                foreach (var GetJFSingleWORow in GetJFSingleWO)
        //                {
        //                    int SettingTime = 0;
        //                    int MachininingTime = 0;
        //                    int LabourTime = 0;

        //                    fileName = "ZCO11" + "_" + startDate.ToString("yyyy -MM-dd").Replace("-", "") + startDate.ToString("HH:mm:ss").Replace(":", "");


        //                    String StartWODate = null;
        //                    String StartWOTime = null;


        //                    StartWODate = Convert.ToDateTime(GetJFSingleWORow.Date).ToString("ddMMyyyy");
        //                    StartWOTime = Convert.ToDateTime(GetJFSingleWORow.Date).ToString("HHmmss");



        //                    string OperatorID = GetJFSingleWORow.OperatorDet;
        //                    if (OperatorID.Contains(","))
        //                    {
        //                        string[] data1 = OperatorID.Split(',');
        //                        OperatorID = data1[0];
        //                    }


        //                    dtNew.Rows.Add(GetJFSingleWORow.WorkOrderNo, Convert.ToInt32(GetJFSingleWORow.OperationNo), GetJFSingleWORow.TargetQty, 0, SettingTime, MachininingTime, LabourTime, StartWODate, StartWOTime, Convert.ToDateTime(GetJFSingleWORow.Time).ToString("ddMMyyyy"), Convert.ToDateTime(GetJFSingleWORow.Time).ToString("HHmmss"), (((2001).ToString() + Convert.ToString(OperatorID)).ToString()));


        //                }


        //                //Push the Data to the Text File
        //                try
        //                {
        //                    fileName = fileName.Replace("-", "");
        //                    fileName = fileName.Replace(":", "");


        //                    finalPathWithFile = System.IO.Path.Combine(folderName, fileName + ".txt");


        //                    #region
        //                    DataView dv1 = new DataView(dtNew);
        //                    dv1.Sort = "OpNo";
        //                    DataTable DTMainFinal = dv1.ToTable();


        //                    for (int ToFileLooper = 0; ToFileLooper < DTMainFinal.Rows.Count; ToFileLooper++)
        //                    {
        //                        writeData = DTMainFinal.Rows[ToFileLooper][0] + "\t" + DTMainFinal.Rows[ToFileLooper][1] + "\t" + DTMainFinal.Rows[ToFileLooper][2] + "\t" + DTMainFinal.Rows[ToFileLooper][3]
        //                             + "\t" + DTMainFinal.Rows[ToFileLooper][4] + "\t" + DTMainFinal.Rows[ToFileLooper][5] + "\t" + DTMainFinal.Rows[ToFileLooper][6] +
        //                             "\t" + DTMainFinal.Rows[ToFileLooper][7] + "\t" + DTMainFinal.Rows[ToFileLooper][8] + "\t" + DTMainFinal.Rows[ToFileLooper][9] + "\t"
        //                             + DTMainFinal.Rows[ToFileLooper][10] + "\t" + DTMainFinal.Rows[ToFileLooper][11] + "\t";



        //                        using (StreamWriter sw = new StreamWriter(@finalPathWithFile, true))
        //                        {
        //                            sw.WriteLine(writeData);


        //                        }

        //                    }
        //                    #endregion

        //                    //getting the report ip address from app string.json
        //                    string reportURL = configuration.GetSection("MySettings").GetSection("ImageUrl").Value;
        //                    string path1 = reportURL + "WOConfirmation/" + fileName + ".txt";

        //                    if (fileName != "")
        //                    {
        //                        obj.isTure = true;
        //                        obj.response = path1;

        //                    }
        //                    else
        //                    {
        //                        obj.isTure = false;
        //                        obj.response = "No Data Found";
        //                    }


        //                }
        //                catch (Exception ex)
        //                {
        //                    log.Error(ex.InnerException.ToString());
        //                    obj.isTure = false;
        //                    obj.response = ResourceResponse.ExceptionMessage;
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                log.Error(ex.InnerException.ToString());
        //                obj.isTure = false;
        //                obj.response = ResourceResponse.ExceptionMessage;
        //            }
        //        }

        //        else if (data.Type == "Split")
        //        {
        //            string finalPathWithFile = "";
        //            DateTime startDate = Convert.ToDateTime(data.Fromdate);
        //            DateTime endDate = Convert.ToDateTime(data.ToDate);
        //            double dateDifference = endDate.Subtract(startDate).TotalDays;
        //            try
        //            {
        //                //var OutputPath = db.Tblpreactorschedule.Where(x => x.IsDeleted == 0 && x.IsStart ==1).FirstOrDefault();

        //                //string fileNameIn = OutputPath.Path + OutputPath.FileFormat + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
        //                var OutputPath = configuration.GetSection("MySettings").GetSection("ImageUrlSave").Value;

        //                string fileNameIn = OutputPath + "WOSplit" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
        //                //var OutputPath = db.Tblpreactorschedule.Where(x => x.IsDeleted == 0 && x.IsStart == 1).FirstOrDefault();

        //                //string fileNameIn = OutputPath.Path + OutputPath.FileFormat + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";

        //                string folderName = OutputPath + "WOSplit/";

        //                if (!Directory.Exists(folderName))
        //                {
        //                    Directory.CreateDirectory(folderName);
        //                }
        //                // Check if file already exists. If yes, delete it.     
        //                if (File.Exists(fileNameIn))
        //                {
        //                    File.Delete(fileNameIn);
        //                }

        //                int startRow = 2;
        //                string fileName = "";
        //                string writeData = "";

        //                DataTable dtNew = new DataTable();
        //                dtNew.Columns.Add("WONo", typeof(string));
        //                dtNew.Columns.Add("OpNo", typeof(int));
        //                dtNew.Columns.Add("Qty", typeof(int));
        //                dtNew.Columns.Add("Scrap", typeof(int));
        //                dtNew.Columns.Add("SettingTime", typeof(int));
        //                dtNew.Columns.Add("MachiningTime", typeof(int));
        //                dtNew.Columns.Add("LabourTime", typeof(int));
        //                dtNew.Columns.Add("StartDate", typeof(string));
        //                dtNew.Columns.Add("StartTime", typeof(string));
        //                dtNew.Columns.Add("EndDate", typeof(string));
        //                dtNew.Columns.Add("EndTime", typeof(string));
        //                dtNew.Columns.Add("OperatorID", typeof(string));

        //                DateTime StartTimeNew = startDate;
        //                DateTime EndTimeNew = endDate;


        //                var wodet = db.Tbllivehmiscreen.Where(m => m.Time > StartTimeNew && m.Time <= EndTimeNew && m.SplitWo == "Yes" && m.IsWorkInProgress == 0 && m.DeliveredQty != 0).ToList();
        //                foreach (var hmirow in wodet)
        //                {

        //                    int SettingTime = 0;
        //                    int MachininingTime = 0;
        //                    int LabourTime = 0;

        //                    String StartWODate = null;
        //                    String StartWOTime = null;

        //                    fileName = "Split_" + "ZCO11" + "_" + startDate.ToString("yyyy -MM-dd").Replace("-", "") + startDate.ToString("HH:mm:ss").Replace(":", "");

        //                    StartWODate = Convert.ToDateTime(hmirow.Date).ToString("ddMMyyyy");
        //                    StartWOTime = Convert.ToDateTime(hmirow.Date).ToString("HHmmss");

        //                    string OperatorID = hmirow.OperatorDet;
        //                    if (OperatorID.Contains(","))
        //                    {
        //                        string[] data1 = OperatorID.Split(',');
        //                        OperatorID = data1[0];
        //                    }
        //                    dtNew.Rows.Add(hmirow.WorkOrderNo, Convert.ToInt32(hmirow.OperationNo), hmirow.DeliveredQty, 0, SettingTime, MachininingTime, LabourTime, StartWODate, StartWOTime, Convert.ToDateTime(hmirow.Time).ToString("ddMMyyyy"), Convert.ToDateTime(hmirow.Time).ToString("HHmmss"), (((2001).ToString() + Convert.ToString(OperatorID)).ToString()));

        //                }

        //                //Push the Data to the Text File
        //                try
        //                {
        //                    fileName = fileName.Replace("-", "");
        //                    fileName = fileName.Replace(":", "");


        //                    finalPathWithFile = System.IO.Path.Combine(folderName, fileName + ".txt");


        //                    #region
        //                    DataView dv1 = new DataView(dtNew);
        //                    dv1.Sort = "OpNo";
        //                    DataTable DTMainFinal = dv1.ToTable();


        //                    for (int ToFileLooper = 0; ToFileLooper < DTMainFinal.Rows.Count; ToFileLooper++)
        //                    {
        //                        writeData = DTMainFinal.Rows[ToFileLooper][0] + "\t" + DTMainFinal.Rows[ToFileLooper][1] + "\t" + DTMainFinal.Rows[ToFileLooper][2] + "\t" + DTMainFinal.Rows[ToFileLooper][3]
        //                             + "\t" + DTMainFinal.Rows[ToFileLooper][4] + "\t" + DTMainFinal.Rows[ToFileLooper][5] + "\t" + DTMainFinal.Rows[ToFileLooper][6] +
        //                             "\t" + DTMainFinal.Rows[ToFileLooper][7] + "\t" + DTMainFinal.Rows[ToFileLooper][8] + "\t" + DTMainFinal.Rows[ToFileLooper][9] + "\t"
        //                             + DTMainFinal.Rows[ToFileLooper][10] + "\t" + DTMainFinal.Rows[ToFileLooper][11] + "\t";



        //                        using (StreamWriter sw = new StreamWriter(@finalPathWithFile, true))
        //                        {
        //                            sw.WriteLine(writeData);


        //                        }

        //                    }
        //                    #endregion
        //                    //getting the report ip address from app string.json
        //                    string reportURL = configuration.GetSection("MySettings").GetSection("ImageUrl").Value;
        //                    string path1 = reportURL + "WOSplit/" + fileName + ".txt";

        //                    if (fileName != "")
        //                    {
        //                        obj.isTure = true;
        //                        obj.response = path1;

        //                    }
        //                    else
        //                    {
        //                        obj.isTure = false;
        //                        obj.response = "No Data Found";
        //                    }

        //                }
        //                catch (Exception ex)
        //                {
        //                    log.Error(ex.InnerException.ToString());
        //                    obj.isTure = false;
        //                    obj.response = ex;
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                log.Error(ex.InnerException.ToString());
        //                obj.isTure = false;
        //                obj.response = ex;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex.InnerException.ToString());
        //        obj.isTure = false;
        //        obj.response = ex;
        //    }
        //    return obj;
        //}

        public string LossLevel(int id)
        {
            string res = "";
            var lossCodesDetails = db.Tbllossescodes.Where(x => x.IsDeleted == 0 && x.LossCodeId == id).Select(x => new { x.LossCodesLevel, x.LossCodesLevel1Id, x.LossCodesLevel2Id, x.LossCodeDesc }).FirstOrDefault();
            if (lossCodesDetails != null)
            {
                int level = lossCodesDetails.LossCodesLevel;
                if (level == 3)
                {
                    int level1Id = (int)lossCodesDetails.LossCodesLevel1Id;
                    int level2Id = (int)lossCodesDetails.LossCodesLevel2Id;
                    res = db.Tbllossescodes.Where(x => x.IsDeleted == 0 && x.LossCodeId == level1Id).Select(x => x.LossCodeDesc).FirstOrDefault() + ",";
                    res += db.Tbllossescodes.Where(x => x.IsDeleted == 0 && x.LossCodeId == level2Id).Select(x => x.LossCodeDesc).FirstOrDefault() + ",";
                    res += lossCodesDetails.LossCodeDesc;
                }
                else if (level == 2)
                {
                    int level1Id = (int)lossCodesDetails.LossCodesLevel1Id;
                    res = db.Tbllossescodes.Where(x => x.IsDeleted == 0 && x.LossCodeId == level1Id).Select(x => x.LossCodeDesc).FirstOrDefault() + ",";
                    res += lossCodesDetails.LossCodeDesc;
                }
                else
                {
                    res = lossCodesDetails.LossCodeDesc;
                }
            }
            return res;
        }

        public static string ExcelColumnFromNumber(int column)
        {

            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        public string GetShift(DateTime dateTime)
        {
            string ShiftValue = "";
            DateTime DateNow = dateTime;
            var ShiftDetails = db.TblshiftMstr.Where(m => m.IsDeleted == 0).ToList();
            foreach (var row in ShiftDetails)
            {
                int ShiftStartHour = row.StartTime.Value.Hours;
                int ShiftEndHour = row.EndTime.Value.Hours;
                int CurrentHour = DateNow.Hour;
                if (CurrentHour >= ShiftStartHour && CurrentHour < ShiftEndHour)
                {
                    ShiftValue = row.ShiftName;
                    break;
                }
            }

            if (ShiftValue == "")
            {
                ShiftValue = "C";
            }

            return ShiftValue;
        }

        #region Suhas report

        //public CommonResponse1 NoComplianceReport(OEEDeckFormat data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();

        //    try
        //    {
        //        #region read date time
        //        DateTime fromDate = DateTime.Now;
        //        try
        //        {
        //            string[] dt = data.fromDate.Split('/');
        //            string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            fromDate = Convert.ToDateTime(frDate);
        //        }
        //        catch
        //        {
        //            fromDate = Convert.ToDateTime(data.fromDate);
        //        }
        //        DateTime toDate = DateTime.Now;
        //        try
        //        {
        //            string[] dt = data.toDate.Split('/');
        //            string torDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            toDate = Convert.ToDateTime(torDate);
        //        }
        //        catch
        //        {
        //            toDate = Convert.ToDateTime(data.toDate).AddHours(24);
        //        }

        //        bool isCondition = false;

        //        if ((data.fromDate != null) && (data.toDate != null))
        //        {
        //            isCondition = true;
        //        }


        //        #endregion

        //        #region 

        //        var getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.IsNormalWc == 0).ToList();
        //        if (data.machineId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.shopId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId && x.IsNormalWc == 0).ToList();
        //        }

        //        #endregion
        //        #region Excel and Stuff
        //        DateTime frda = DateTime.Now;

        //        //getting the connection string from app string.json
        //        string ImageUrlSave = configuration.GetSection("MySettings").GetSection("ImageUrlSave").Value;
        //        string ImageUrl = configuration.GetSection("MySettings").GetSection("ImageUrl").Value;


        //        String FileDir = ImageUrlSave + "\\" + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //        String retrivalPath = ImageUrl + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //        FileInfo newFile = new FileInfo(FileDir);

        //        if (newFile.Exists)
        //        {
        //            try
        //            {
        //                newFile.Delete();  // ensures we create a new workbook
        //                newFile = new FileInfo(FileDir);
        //            }
        //            catch (Exception ex)
        //            {
        //                //ErrorLog.SendErrorToDB(ex);
        //                obj.response = ResourceResponse.ExceptionMessage; ;
        //            }
        //        }

        //        FileInfo templateFile = new FileInfo("C:\\TataReport\\NewTemplates\\ComplianceReportTemplate.xlsx");
        //        ExcelPackage templatep = new ExcelPackage(templateFile);
        //        ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];
        //        ExcelWorksheet Templatews1 = templatep.Workbook.Worksheets[1];
        //        ExcelWorksheet Templatews2 = templatep.Workbook.Worksheets[2];
        //        ExcelWorksheet Templatews3 = templatep.Workbook.Worksheets[3];
        //        ExcelWorksheet Templatews4 = templatep.Workbook.Worksheets[4];
        //        //Using the File for generation and populating it
        //        ExcelPackage p = null;
        //        p = new ExcelPackage(newFile);
        //        ExcelWorksheet worksheet = null;
        //        ExcelWorksheet worksheet1 = null;
        //        ExcelWorksheet worksheet2 = null;
        //        ExcelWorksheet worksheet3 = null;
        //        ExcelWorksheet worksheet4 = null;

        //        //Creating the WorkSheet for populating
        //        try
        //        {
        //            worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet4 = p.Workbook.Worksheets.Add("Graph", Templatews4);
        //        }
        //        catch (Exception ex) { }

        //        if (worksheet == null)
        //        {
        //            worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);

        //        }
        //        if (worksheet1 == null)
        //        {
        //            worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
        //        }
        //        if (worksheet2 == null)
        //        {
        //            worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
        //        }
        //        if (worksheet3 == null)
        //        {
        //            worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
        //        }
        //        if (worksheet4 == null)
        //        {
        //            worksheet4 = p.Workbook.Worksheets.Add("NoCompliance", Templatews4);
        //        }

        //        int sheetcount = p.Workbook.Worksheets.Count;
        //        //p.Workbook.Worksheets.MoveToStart(sheetcount);
        //        //worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //        //worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //        #endregion

        //        try
        //        {

        //            #region Stored procedure calling and to check what is the data to be sent for no date search
        //            string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
        //            //SqlConnection SqlConnection = new SqlConnection(connectionString);
        //            #endregion


        //            DataTable dataTable = new DataTable();
        //            DataTable dataTable1 = new DataTable();
        //            DataTable dataTable2 = new DataTable();
        //            DataTable dataTable3 = new DataTable();

        //            try
        //            {
        //                #region fill data in excel

        //                #region call sp
        //                #region Loss Register Report

        //                var chart = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);
        //                var chart1 = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart1", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);

        //                //var test = getMachineDetails.Take(18);

        //                foreach (var item in getMachineDetails)
        //                {
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportLossRegister", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable);
        //                            }
        //                        }
        //                    }

        //                    int count = dataTable.Rows.Count;
        //                    if (count > 0)
        //                    {
        //                        worksheet.Cells["C3"].Value = data.fromDate;
        //                        worksheet.Cells["E3"].Value = data.toDate;
        //                        worksheet.Cells["B5"].LoadFromDataTable(dataTable, true);
        //                    }
        //                    #endregion

        //                    #region Un Assigned WO Report
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportUnAssignedWO", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable1);
        //                            }
        //                        }
        //                    }
        //                    int count1 = dataTable1.Rows.Count;
        //                    if (count1 > 0)
        //                    {
        //                        worksheet1.Cells["B5"].LoadFromDataTable(dataTable1, true);
        //                    }
        //                    #endregion

        //                    #region Loss Code Report No Login
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportLossCodeNoLogin", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable2);
        //                            }
        //                        }
        //                    }
        //                    int count2 = dataTable2.Rows.Count;
        //                    if (count2 > 0)
        //                    {
        //                        worksheet2.Cells["B5"].LoadFromDataTable(dataTable2, true);
        //                    }
        //                    #endregion

        //                    #region No Compliance Overall Report
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportNocompliance", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable3);
        //                            }
        //                        }
        //                    }

        //                    int count3 = dataTable3.Rows.Count;
        //                    if (count3 > 0)
        //                    {
        //                        worksheet3.Cells["B4"].LoadFromDataTable(dataTable3, true);
        //                    }

        //                    #endregion

        //                    #region Graph

        //                    #region Chart Data

        //                    int columnNumber = 1;
        //                    int alpha = 1;
        //                    worksheet4.Cells["A1"].Value = "Date";
        //                    worksheet4.Cells["A2"].Value = "NoCode";
        //                    worksheet4.Cells["A3"].Value = "UnAssignedWO";
        //                    worksheet4.Cells["A4"].Value = "NoLogin";

        //                    worksheet4.Cells["A5"].Value = "NoCodePercentage";
        //                    worksheet4.Cells["A6"].Value = "UnAssignedWOPercentage";
        //                    worksheet4.Cells["A7"].Value = "NoLoginPercentage";
        //                    string finalAplha = "";
        //                    for (int i = 0; i < count3; i++)
        //                    {
        //                        //string alphabet = CommonFunction.GetExcelColumnName(i + 2);
        //                        string alphabet = GetExcelColumnName(i + 2);
        //                        string date = dataTable3.Rows[i][1].ToString();
        //                        string noCode = dataTable3.Rows[i][4].ToString();
        //                        string unAssignedWO = dataTable3.Rows[i][5].ToString();
        //                        string noLogin = dataTable3.Rows[i][6].ToString();
        //                        string noCodePercentage = dataTable3.Rows[i][10].ToString();
        //                        string unAssignedWOPercentage = dataTable3.Rows[i][11].ToString();
        //                        string noLoginPercentage = dataTable3.Rows[i][12].ToString();

        //                        #region Excel Graph Data
        //                        worksheet4.Cells[alphabet + 1].Value = date;
        //                        worksheet4.Cells[alphabet + 5].Value = date;
        //                        decimal zeroValue = 0;
        //                        #region Chart 1


        //                        if (noCode != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 2].Value = Convert.ToInt32(noCode);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 2].Value = zeroValue;
        //                        }
        //                        if (unAssignedWO != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 3].Value = Convert.ToInt32(unAssignedWO);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 3].Value = zeroValue;
        //                        }
        //                        if (noLogin != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 4].Value = Convert.ToInt32(noLogin);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 4].Value = zeroValue;
        //                        }

        //                        #endregion
        //                        #region Chart 2

        //                        if (noCodePercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 5].Value = Convert.ToDecimal(noCodePercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 5].Value = zeroValue;
        //                        }
        //                        if (unAssignedWOPercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 6].Value = Convert.ToDecimal(unAssignedWOPercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 6].Value = zeroValue;
        //                        }
        //                        if (noLoginPercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 7].Value = Convert.ToDecimal(noLoginPercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 7].Value = zeroValue;
        //                        }


        //                        #endregion

        //                        #endregion
        //                        //finalAplha = CommonFunction.GetExcelColumnName(alpha + 1);
        //                        finalAplha = GetExcelColumnName(alpha + 1);
        //                        alpha++;
        //                        columnNumber++;
        //                    }

        //                    worksheet4.Cells[2, count3 + 2].Style.Numberformat.Format = "##0.0";
        //                    worksheet4.Cells[2, count3 + 2].Style.Numberformat.Format = "##0.0";
        //                    #endregion

        //                    #region Chart Graph
        //                    try
        //                    {
        //                        for (int i = 2; i <= 4; i++)
        //                        {
        //                            //string alphabet = CommonFunction.GetExcelColumnName(alpha + 1);
        //                            string alphabet = GetExcelColumnName(alpha + 1);
        //                            #region chart 1
        //                            chart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        //                            chart.Legend.Add();
        //                            //chart.SetPosition(SerialNo + 8, 0, 1, 0);
        //                            chart.SetPosition(0, 0, 0, 0);
        //                            chart.SetSize(count3 * 100, 500);
        //                            chart.DataLabel.ShowValue = true;

        //                            var series1 = chart.Series.Add($"B{i}:" + finalAplha + $"{i}", "B1:" + finalAplha + "1");
        //                            series1.Header = worksheet4.Cells[$"A{i}"].Value.ToString();

        //                            chart.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
        //                            chart.Title.Text = "No Compliance Value Report";
        //                            chart.XAxis.Title.Text = "Day";
        //                            chart.YAxis.Title.Text = "Value";
        //                            #endregion
        //                            #region chart 2
        //                            chart1.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        //                            chart1.Legend.Add();
        //                            //chart.SetPosition(SerialNo + 8, 0, 1, 0);
        //                            chart1.SetPosition(26, 0, 0, 0);
        //                            chart1.SetSize(count3 * 100, 500);
        //                            chart1.DataLabel.ShowValue = true;

        //                            int j = i + 3;
        //                            var series2 = chart1.Series.Add($"B{j}:" + finalAplha + $"{j}", "B1:" + finalAplha + "1");

        //                            series2.Header = worksheet4.Cells[$"A{j}"].Value.ToString();

        //                            chart1.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
        //                            chart1.Title.Text = "No Compliance Percentage Report";
        //                            chart1.XAxis.Title.Text = "Day";
        //                            chart1.YAxis.Title.Text = "Percentage";
        //                            #endregion

        //                        }

        //                    }

        //                    catch (Exception ex)
        //                    {

        //                    }
        //                    #endregion



        //                    #endregion

        //                    #endregion

        //                    #endregion

        //                    #region Save and Download

        //                    p.Save();

        //                    //Downloding Excel
        //                    string path = retrivalPath;

        //                    #endregion

        //                    obj.isStatus = true;
        //                    obj.response = path;
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
        //                obj.response = ResourceResponse.ExceptionMessage;
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
        //            //ErrorLog.SendErrorToDB(e);
        //            obj.response = ResourceResponse.ExceptionMessage; ;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        obj.isStatus = false;
        //    }

        //    return obj;
        //}


        //public CommonResponse1 NoComplianceReport(OEEDeckFormat data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();

        //    try
        //    {
        //        #region read date time
        //        DateTime fromDate = DateTime.Now;
        //        try
        //        {
        //            string[] dt = data.fromDate.Split('/');
        //            string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            fromDate = Convert.ToDateTime(frDate);
        //        }
        //        catch
        //        {
        //            fromDate = Convert.ToDateTime(data.fromDate);
        //        }
        //        DateTime toDate = DateTime.Now;
        //        try
        //        {
        //            string[] dt = data.toDate.Split('/');
        //            string torDate = dt[2] + '-' + dt[1] + '-' + dt[0];
        //            toDate = Convert.ToDateTime(torDate);
        //        }
        //        catch
        //        {
        //            toDate = Convert.ToDateTime(data.toDate).AddHours(24);
        //        }

        //        bool isCondition = false;

        //        if ((data.fromDate != null) && (data.toDate != null))
        //        {
        //            isCondition = true;
        //        }


        //        #endregion

        //        #region 

        //        var getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.IsNormalWc == 0).ToList();
        //        if (data.machineId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.cellId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.shopId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId && x.IsNormalWc == 0).ToList();
        //        }
        //        else if (data.plantId != 0)
        //        {
        //            getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId && x.IsNormalWc == 0).ToList();
        //        }

        //        #endregion
        //        #region Excel and Stuff
        //        DateTime frda = DateTime.Now;

        //        //getting the connection string from app string.json
        //        string ImageUrlSave = configuration.GetSection("AppSettings").GetSection("ImageUrlSave").Value;
        //        string ImageUrl = configuration.GetSection("AppSettings").GetSection("ImageUrl").Value;


        //        String FileDir = ImageUrlSave + "\\" + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //        String retrivalPath = ImageUrl + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
        //        FileInfo newFile = new FileInfo(FileDir);

        //        if (newFile.Exists)
        //        {
        //            try
        //            {
        //                newFile.Delete();  // ensures we create a new workbook
        //                newFile = new FileInfo(FileDir);
        //            }
        //            catch (Exception ex)
        //            {
        //                //ErrorLog.SendErrorToDB(ex);
        //                obj.response = ResourceResponse.ExceptionMessage; ;
        //            }
        //        }

        //        FileInfo templateFile = new FileInfo("C:\\TataReport\\NewTemplates\\ComplianceReportTemplate.xlsx");
        //        ExcelPackage templatep = new ExcelPackage(templateFile);
        //        ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];
        //        ExcelWorksheet Templatews1 = templatep.Workbook.Worksheets[1];
        //        ExcelWorksheet Templatews2 = templatep.Workbook.Worksheets[2];
        //        ExcelWorksheet Templatews3 = templatep.Workbook.Worksheets[3];
        //        ExcelWorksheet Templatews4 = templatep.Workbook.Worksheets[4];
        //        //Using the File for generation and populating it
        //        ExcelPackage p = null;
        //        p = new ExcelPackage(newFile);
        //        ExcelWorksheet worksheet = null;
        //        ExcelWorksheet worksheet1 = null;
        //        ExcelWorksheet worksheet2 = null;
        //        ExcelWorksheet worksheet3 = null;
        //        ExcelWorksheet worksheet4 = null;

        //        //Creating the WorkSheet for populating
        //        try
        //        {
        //            worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
        //        }
        //        catch (Exception ex) { }

        //        try
        //        {
        //            worksheet4 = p.Workbook.Worksheets.Add("Graph", Templatews4);
        //        }
        //        catch (Exception ex) { }

        //        if (worksheet == null)
        //        {
        //            worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);

        //        }
        //        if (worksheet1 == null)
        //        {
        //            worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
        //        }
        //        if (worksheet2 == null)
        //        {
        //            worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
        //        }
        //        if (worksheet3 == null)
        //        {
        //            worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
        //        }
        //        if (worksheet4 == null)
        //        {
        //            worksheet4 = p.Workbook.Worksheets.Add("NoCompliance", Templatews4);
        //        }

        //        int sheetcount = p.Workbook.Worksheets.Count;
        //        //p.Workbook.Worksheets.MoveToStart(sheetcount);
        //        //worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //        //worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //        #endregion

        //        try
        //        {

        //            #region Stored procedure calling and to check what is the data to be sent for no date search
        //            string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
        //            //SqlConnection SqlConnection = new SqlConnection(connectionString);
        //            #endregion


        //            DataTable dataTable = new DataTable();
        //            DataTable dataTable1 = new DataTable();
        //            DataTable dataTable2 = new DataTable();
        //            DataTable dataTable3 = new DataTable();


        //            DataTable dataTable4 = new DataTable();
        //            DataTable dataTable5 = new DataTable();
        //            DataTable dataTable6 = new DataTable();
        //            DataTable dataTable7 = new DataTable();
        //            try
        //            {
        //                #region fill data in excel

        //                #region call sp
        //                #region Loss Register Report

        //                var chart = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);
        //                var chart1 = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart1", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);

        //                //var test = getMachineDetails.Take(18);

        //                foreach (var item in getMachineDetails)
        //                {
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportLossRegister", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable);
        //                            }
        //                        }

        //                    }

        //                    dataTable4.Merge(dataTable);
        //                    for (int t = 0; t < dataTable4.Rows.Count; t++)
        //                    {
        //                        //int l = t;
        //                        //dataTable4.Rows[0][0] = l + 1;
        //                        dataTable4.Rows[t][0] = t + 1;
        //                    }

        //                    int count = dataTable4.Rows.Count;
        //                    if (count > 0)
        //                    {
        //                        worksheet.Cells["3"].Value = data.fromDate;
        //                        worksheet.Cells["F3"].Value = data.toDate;
        //                        worksheet.Cells["B5"].LoadFromDataTable(dataTable4, true);
        //                    }
        //                    #endregion

        //                    #region Un Assigned WO Report
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportUnAssignedWO", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable1);
        //                            }
        //                        }
        //                    }

        //                    dataTable5.Merge(dataTable1);
        //                    for (int t = 0; t < dataTable5.Rows.Count; t++)
        //                    {
        //                        //int l = t;
        //                        //dataTable4.Rows[0][0] = l + 1;
        //                        dataTable5.Rows[t][0] = t + 1;
        //                    }

        //                    int count1 = dataTable5.Rows.Count;
        //                    if (count1 > 0)
        //                    {
        //                        worksheet1.Cells["B4"].LoadFromDataTable(dataTable5, true);
        //                    }
        //                    #endregion

        //                    #region Loss Code Report No Login
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportLossCodeNoLogin", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlCmd.CommandTimeout = 300;
        //                                sqlAdapter.Fill(dataTable2);
        //                            }
        //                        }
        //                    }

        //                    dataTable6.Merge(dataTable2);
        //                    for (int t = 0; t < dataTable6.Rows.Count; t++)
        //                    {
        //                        //int l = t;
        //                        //dataTable4.Rows[0][0] = l + 1;
        //                        dataTable6.Rows[t][0] = t + 1;
        //                    }

        //                    int count2 = dataTable6.Rows.Count;
        //                    if (count2 > 0)
        //                    {
        //                        worksheet2.Cells["B4"].LoadFromDataTable(dataTable6, true);
        //                    }
        //                    #endregion

        //                    #region No Compliance Overall Report
        //                    using (SqlConnection sqlConn = new SqlConnection(connectionString))
        //                    {
        //                        using (SqlCommand sqlCmd = new SqlCommand("ReportNocompliance", sqlConn))
        //                        {
        //                            sqlCmd.CommandType = CommandType.StoredProcedure;
        //                            sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
        //                            sqlCmd.Parameters.AddWithValue("@toDate", toDate);
        //                            sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
        //                            sqlConn.Open();
        //                            using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
        //                            {
        //                                sqlAdapter.Fill(dataTable3);
        //                            }
        //                        }
        //                    }

        //                    dataTable7.Merge(dataTable3);
        //                    for (int t = 0; t < dataTable7.Rows.Count; t++)
        //                    {
        //                        //int l = t;
        //                        //dataTable4.Rows[0][0] = l + 1;
        //                        dataTable7.Rows[t][0] = t + 1;
        //                    }

        //                    int count4 = dataTable7.Rows.Count;
        //                    int count3 = dataTable3.Rows.Count;
        //                    if (count4 > 0)
        //                    {
        //                        worksheet3.Cells["B4"].LoadFromDataTable(dataTable7, true);
        //                    }

        //                    #endregion

        //                    #region Graph

        //                    #region Chart Data

        //                    int columnNumber = 1;
        //                    int alpha = 1;
        //                    worksheet4.Cells["A1"].Value = "Date";
        //                    worksheet4.Cells["A2"].Value = "NoCode";
        //                    worksheet4.Cells["A3"].Value = "UnAssignedWO";
        //                    worksheet4.Cells["A4"].Value = "NoLogin";

        //                    worksheet4.Cells["A5"].Value = "NoCodePercentage";
        //                    worksheet4.Cells["A6"].Value = "UnAssignedWOPercentage";
        //                    worksheet4.Cells["A7"].Value = "NoLoginPercentage";
        //                    string finalAplha = "";
        //                    for (int i = 0; i < count3; i++)
        //                    {
        //                        //string alphabet = CommonFunction.GetExcelColumnName(i + 2);
        //                        string alphabet = GetExcelColumnName(i + 2);
        //                        string date = dataTable3.Rows[i][1].ToString();
        //                        string noCode = dataTable3.Rows[i][4].ToString();
        //                        string unAssignedWO = dataTable3.Rows[i][5].ToString();
        //                        string noLogin = dataTable3.Rows[i][6].ToString();
        //                        string noCodePercentage = dataTable3.Rows[i][10].ToString();
        //                        string unAssignedWOPercentage = dataTable3.Rows[i][11].ToString();
        //                        string noLoginPercentage = dataTable3.Rows[i][12].ToString();

        //                        #region Excel Graph Data
        //                        worksheet4.Cells[alphabet + 1].Value = date;
        //                        worksheet4.Cells[alphabet + 5].Value = date;
        //                        decimal zeroValue = 0;
        //                        #region Chart 1


        //                        if (noCode != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 2].Value = Convert.ToInt32(noCode);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 2].Value = zeroValue;
        //                        }
        //                        if (unAssignedWO != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 3].Value = Convert.ToInt32(unAssignedWO);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 3].Value = zeroValue;
        //                        }
        //                        if (noLogin != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 4].Value = Convert.ToInt32(noLogin);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 4].Value = zeroValue;
        //                        }

        //                        #endregion
        //                        #region Chart 2

        //                        if (noCodePercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 5].Value = Convert.ToDecimal(noCodePercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 5].Value = zeroValue;
        //                        }
        //                        if (unAssignedWOPercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 6].Value = Convert.ToDecimal(unAssignedWOPercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 6].Value = zeroValue;
        //                        }
        //                        if (noLoginPercentage != "")
        //                        {
        //                            worksheet4.Cells[alphabet + 7].Value = Convert.ToDecimal(noLoginPercentage);
        //                        }
        //                        else
        //                        {
        //                            worksheet4.Cells[alphabet + 7].Value = zeroValue;
        //                        }


        //                        #endregion

        //                        #endregion
        //                        //finalAplha = CommonFunction.GetExcelColumnName(alpha + 1);
        //                        finalAplha = GetExcelColumnName(alpha + 1);
        //                        alpha++;
        //                        columnNumber++;
        //                    }

        //                    ExcelRange rng = worksheet4.Cells["B1:B"+ count3 + 1];
        //                    rng.Style.Numberformat.Format = "#";
        //                    //worksheet4.Cells[2, count3 + 2].Style.Numberformat.Format = "##0.0";
        //                    #endregion

        //                    #region Chart Graph
        //                    try
        //                    {
        //                        for (int i = 2; i <= 4; i++)
        //                        {
        //                            //string alphabet = CommonFunction.GetExcelColumnName(alpha + 1);
        //                            string alphabet = GetExcelColumnName(alpha + 1);
        //                            #region chart 1
        //                            chart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        //                            chart.Legend.Add();
        //                            //chart.SetPosition(SerialNo + 8, 0, 1, 0);
        //                            chart.SetPosition(0, 0, 0, 0);
        //                            chart.SetSize(count3 * 100, 500);
        //                            chart.DataLabel.ShowValue = true;

        //                            var series1 = chart.Series.Add($"B{i}:" + finalAplha + $"{i}", "B1:" + finalAplha + "1");
        //                            series1.Header = worksheet4.Cells[$"A{i}"].Value.ToString();

        //                            chart.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
        //                            chart.Title.Text = "No Compliance Value Report";
        //                            chart.XAxis.Title.Text = "Day";
        //                            chart.YAxis.Title.Text = "Value";
        //                            #endregion
        //                            #region chart 2
        //                            chart1.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
        //                            chart1.Legend.Add();
        //                            //chart.SetPosition(SerialNo + 8, 0, 1, 0);
        //                            chart1.SetPosition(26, 0, 0, 0);
        //                            chart1.SetSize(count3 * 100, 500);
        //                            chart1.DataLabel.ShowValue = true;

        //                            int j = i + 3;
        //                            var series2 = chart1.Series.Add($"B{j}:" + finalAplha + $"{j}", "B1:" + finalAplha + "1");

        //                            series2.Header = worksheet4.Cells[$"A{j}"].Value.ToString();

        //                            chart1.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
        //                            chart1.Title.Text = "No Compliance Percentage Report";
        //                            chart1.XAxis.Title.Text = "Day";
        //                            chart1.YAxis.Title.Text = "Percentage";
        //                            #endregion

        //                        }

        //                    }

        //                    catch (Exception ex)
        //                    {

        //                    }
        //                    #endregion



        //                    #endregion

        //                    #endregion

        //                    #endregion

        //                    #region Save and Download

        //                    p.Save();

        //                    //Downloding Excel
        //                    string path = retrivalPath;

        //                    #endregion

        //                    obj.isStatus = true;
        //                    obj.response = path;
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
        //                obj.response = ResourceResponse.ExceptionMessage;
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
        //            //ErrorLog.SendErrorToDB(e);
        //            obj.response = ResourceResponse.ExceptionMessage; ;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.response = ResourceResponse.ExceptionMessage;
        //        obj.isStatus = false;
        //    }

        //    return obj;
        //}

        public CommonResponse1 NoComplianceReport(OEEDeckFormat data)
        {
            CommonResponse1 obj = new CommonResponse1();

            try
            {
                #region read date time
                DateTime fromDate = DateTime.Now;
                try
                {
                    string[] dt = data.fromDate.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    fromDate = Convert.ToDateTime(frDate);
                }
                catch
                {
                    fromDate = Convert.ToDateTime(data.fromDate);
                }
                DateTime toDate = DateTime.Now;
                try
                {
                    string[] dt = data.toDate.Split('/');
                    string torDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    toDate = Convert.ToDateTime(torDate);
                }
                catch
                {
                    toDate = Convert.ToDateTime(data.toDate).AddHours(24);
                }


                bool isCondition = false;


                if ((data.fromDate != null) && (data.toDate != null))
                {
                    isCondition = true;
                }


                #endregion

                #region 

                var getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.IsNormalWc == 0).ToList();
                if (data.machineId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId && x.IsNormalWc == 0).ToList();
                }
                else if (data.cellId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId && x.IsNormalWc == 0).ToList();
                }
                else if (data.shopId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId && x.IsNormalWc == 0).ToList();
                }
                else if (data.plantId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId && x.IsNormalWc == 0).ToList();
                }

                #endregion
                #region Excel and Stuff
                DateTime frda = DateTime.Now;

                //getting the connection string from app string.json
                string ImageUrlSave = configuration.GetSection("AppSettings").GetSection("ImageUrlSave").Value;
                string ImageUrl = configuration.GetSection("AppSettings").GetSection("ImageUrl").Value;


                String FileDir = ImageUrlSave + "\\" + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                String retrivalPath = ImageUrl + "No Compliance Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                FileInfo newFile = new FileInfo(FileDir);

                if (newFile.Exists)
                {
                    try
                    {
                        newFile.Delete();  // ensures we create a new workbook
                        newFile = new FileInfo(FileDir);
                    }

                    catch (Exception ex)

                    {
                        //ErrorLog.SendErrorToDB(ex);
                        obj.response = ResourceResponse.ExceptionMessage; ;
                    }
                }

                FileInfo templateFile = new FileInfo("C:\\TataReport\\NewTemplates\\ComplianceReportTemplate.xlsx");
                ExcelPackage templatep = new ExcelPackage(templateFile);
                ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];
                ExcelWorksheet Templatews1 = templatep.Workbook.Worksheets[1];
                ExcelWorksheet Templatews2 = templatep.Workbook.Worksheets[2];
                ExcelWorksheet Templatews3 = templatep.Workbook.Worksheets[3];
                ExcelWorksheet Templatews4 = templatep.Workbook.Worksheets[4];
                //Using the File for generation and populating it
                ExcelPackage p = null;
                p = new ExcelPackage(newFile);
                ExcelWorksheet worksheet = null;
                ExcelWorksheet worksheet1 = null;
                ExcelWorksheet worksheet2 = null;
                ExcelWorksheet worksheet3 = null;
                ExcelWorksheet worksheet4 = null;

                //Creating the WorkSheet for populating
                try
                {
                    worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);
                }

                catch (Exception ex) { }

                try
                {
                    worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
                }

                catch (Exception ex) { }


                try
                {
                    worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
                }

                catch (Exception ex) { }


                try
                {
                    worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
                }

                catch (Exception ex) { }


                try
                {
                    worksheet4 = p.Workbook.Worksheets.Add("Graph", Templatews4);
                }

                catch (Exception ex) { }


                if (worksheet == null)
                {
                    worksheet = p.Workbook.Worksheets.Add("LossRegister", Templatews);

                }
                if (worksheet1 == null)
                {
                    worksheet1 = p.Workbook.Worksheets.Add("UnAssignedWO", Templatews1);
                }
                if (worksheet2 == null)
                {
                    worksheet2 = p.Workbook.Worksheets.Add("NoLogin", Templatews2);
                }
                if (worksheet3 == null)
                {
                    worksheet3 = p.Workbook.Worksheets.Add("NoCompliance", Templatews3);
                }
                if (worksheet4 == null)
                {
                    worksheet4 = p.Workbook.Worksheets.Add("NoCompliance", Templatews4);
                }

                int sheetcount = p.Workbook.Worksheets.Count;
                //p.Workbook.Worksheets.MoveToStart(sheetcount);
                //worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                #endregion

                try
                {

                    #region Stored procedure calling and to check what is the data to be sent for no date search
                    string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                    //SqlConnection SqlConnection = new SqlConnection(connectionString);
                    #endregion


                    DataTable dataTable = new DataTable();
                    DataTable dataTable1 = new DataTable();
                    DataTable dataTable2 = new DataTable();
                    DataTable dataTable3 = new DataTable();


                    DataTable dataTable4 = new DataTable();
                    DataTable dataTable5 = new DataTable();
                    DataTable dataTable6 = new DataTable();
                    DataTable dataTable7 = new DataTable();
                    try
                    {
                        #region fill data in excel

                        #region call sp
                        #region Loss Register Report

                        var chart = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);
                        var chart1 = (OfficeOpenXml.Drawing.Chart.ExcelBarChart)worksheet4.Drawings.AddChart("ContiARSChart1", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D);

                        //var test = getMachineDetails.Take(18);
                        List<HMIDetails> HMIDetailsList = new List<HMIDetails>();
                        foreach (var item in getMachineDetails)
                        {
                            using (SqlConnection sqlConn = new SqlConnection(connectionString))
                            {
                                using (SqlCommand sqlCmd = new SqlCommand("ReportLossRegister", sqlConn))
                                {
                                    sqlCmd.CommandType = CommandType.StoredProcedure;
                                    sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                    sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                    sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
                                    sqlConn.Open();
                                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                    {
                                        sqlAdapter.Fill(dataTable);
                                    }
                                }

                            }

                            dataTable4.Merge(dataTable);
                            for (int t = 0; t < dataTable4.Rows.Count; t++)
                            {
                                //int l = t;
                                //dataTable4.Rows[0][0] = l + 1;
                                dataTable4.Rows[t][0] = t + 1;
                            }

                            int count = dataTable4.Rows.Count;
                            if (count > 0)
                            {
                                worksheet.Cells["3"].Value = data.fromDate;
                                worksheet.Cells["F3"].Value = data.toDate;
                                worksheet.Cells["B5"].LoadFromDataTable(dataTable4, true);
                            }
                            #endregion

                            #region Un Assigned WO Report
                            using (SqlConnection sqlConn = new SqlConnection(connectionString))
                            {
                                using (SqlCommand sqlCmd = new SqlCommand("ReportUnAssignedWO", sqlConn))
                                {
                                    sqlCmd.CommandType = CommandType.StoredProcedure;
                                    sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                    sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                    sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
                                    sqlConn.Open();
                                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                    {
                                        sqlAdapter.Fill(dataTable1);
                                    }
                                }
                            }

                            if (dataTable1.Rows.Count > 0)
                            {
                                dataTable5.Merge(dataTable1);
                                for (int t = 0; t < dataTable5.Rows.Count; t++)
                                {
                                    //int l = t;
                                    //dataTable4.Rows[0][0] = l + 1;
                                    dataTable5.Rows[t][0] = t + 1;
                                }

                                int count1 = dataTable5.Rows.Count;
                                if (count1 > 0)
                                {
                                    worksheet1.Cells["B4"].LoadFromDataTable(dataTable5, true);
                                }
                            }
                            else
                            {
                                int StartRow = 5;
                                int SlNo = 1;
                                DateTime temp = Convert.ToDateTime(fromDate);
                                DateTime temp1 = Convert.ToDateTime(toDate);
                                string dayTimingStartTime = Convert.ToString(db.Tbldaytiming.Where(x => x.IsDeleted == 0).Select(x => x.StartTime).FirstOrDefault());
                                string dayStart = temp.ToString("yyyy-MM-dd");
                                string dayEnd = temp1.ToString("yyyy-MM-dd");
                                var machineName = db.Tblmachinedetails.Where(m => m.MachineId == item.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                                var plantName = db.Tblplant.Where(m => m.PlantId == db.Tblmachinedetails.Where(n => n.MachineId == item.MachineId).Select(n => n.PlantId).FirstOrDefault()).Select(m => m.PlantName).FirstOrDefault();
                                var shopName = db.Tblshop.Where(m => m.PlantId == db.Tblmachinedetails.Where(n => n.MachineId == item.MachineId).Select(n => n.ShopId).FirstOrDefault()).Select(m => m.ShopName).FirstOrDefault();
                                var cellName = db.Tblcell.Where(m => m.CellId == db.Tblmachinedetails.Where(n => n.MachineId == item.MachineId).Select(n => n.CellId).FirstOrDefault()).Select(m => m.CellName).FirstOrDefault();
                                var hmiData = db.Tblhmiscreen.Where(x => x.MachineId == item.MachineId && Convert.ToDateTime(x.CorrectedDate) >= temp && Convert.ToDateTime(x.CorrectedDate) <= temp1 && (x.Status == 2 || x.Status == 1) && (x.IsWorkInProgress == 1 || x.IsWorkInProgress == 0)).Select(x => new { x.Date, x.Time, x.CorrectedDate }).OrderBy(x => x.Date).ToList();
                                if (hmiData.Count > 0)
                                {
                                    for (int j = 0; j <= hmiData.Count(); j++)
                                    {
                                        HMIDetails HMIDetails = new HMIDetails();
                                        HMIDetails.plantName = plantName;
                                        HMIDetails.shopName = shopName;
                                        HMIDetails.cellName = cellName;
                                        HMIDetails.machineName = machineName;
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
                                                HMIDetails.startDateTime = dayStart;
                                                HMIDetails.endDateTime = dayStartTbl;
                                                HMIDetails.correctedDate = hmiData[j].CorrectedDate;
                                                HMIDetails.machineId = item.MachineId;
                                                HMIDetails.duration = getMinutes;
                                            }
                                        }
                                        else if (j == hmiData.Count())// ending time row
                                        {
                                            string previousEndTime = Convert.ToString(hmiData[j - 1].Time);
                                            string tempStartTime = Convert.ToDateTime(temp).AddDays(1).ToString("yyyy-MM-dd") + " " + dayTimingStartTime;
                                            string presentStartTime = Convert.ToDateTime(tempStartTime).AddSeconds(-1).ToString("yyyy-MM-dd HH:mm:ss");
                                            double getMinutes = Convert.ToDateTime(presentStartTime).Subtract(Convert.ToDateTime(previousEndTime)).TotalMinutes;
                                            if (presentStartTime == previousEndTime)
                                            {
                                                // nothing to do
                                            }
                                            else if (getMinutes >= 1)
                                            {
                                                HMIDetails.startDateTime = previousEndTime;
                                                HMIDetails.endDateTime = presentStartTime;
                                                HMIDetails.correctedDate = hmiData[j - 1].CorrectedDate;
                                                HMIDetails.machineId = item.MachineId;
                                                HMIDetails.duration = getMinutes;
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
                                                HMIDetails.startDateTime = previousEndTime;
                                                HMIDetails.endDateTime = presentStartTime;
                                                HMIDetails.correctedDate = hmiData[j].CorrectedDate;
                                                HMIDetails.machineId = item.MachineId;
                                                HMIDetails.duration = getMinutes;
                                            }
                                        }
                                        HMIDetailsList.Add(HMIDetails);
                                    }
                                    foreach (var items in HMIDetailsList)
                                    {
                                        worksheet1.Cells["B" + StartRow].Value = SlNo;
                                        worksheet1.Cells["C" + StartRow].Value = items.plantName;
                                        worksheet1.Cells["D" + StartRow].Value = items.shopName;
                                        worksheet1.Cells["E" + StartRow].Value = items.cellName;
                                        worksheet1.Cells["F" + StartRow].Value = items.machineName;
                                        worksheet1.Cells["G" + StartRow].Value = items.correctedDate;
                                        worksheet1.Cells["H" + StartRow].Value = items.startDateTime;
                                        worksheet1.Cells["I" + StartRow].Value = items.endDateTime;
                                        worksheet1.Cells["J" + StartRow].Value = items.duration;
                                        SlNo++;
                                        StartRow++;
                                    }
                                }
                            }
                            #endregion

                            #region Loss Code Report No Login
                            using (SqlConnection sqlConn = new SqlConnection(connectionString))
                            {
                                using (SqlCommand sqlCmd = new SqlCommand("ReportLossCodeNoLogin", sqlConn))
                                {
                                    sqlCmd.CommandType = CommandType.StoredProcedure;
                                    sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                    sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                    sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
                                    sqlConn.Open();
                                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                    {
                                        sqlCmd.CommandTimeout = 300;
                                        sqlAdapter.Fill(dataTable2);
                                    }
                                }
                            }

                            dataTable6.Merge(dataTable2);
                            for (int t = 0; t < dataTable6.Rows.Count; t++)
                            {
                                //int l = t;
                                //dataTable4.Rows[0][0] = l + 1;
                                dataTable6.Rows[t][0] = t + 1;
                            }

                            int count2 = dataTable6.Rows.Count;
                            if (count2 > 0)
                            {
                                worksheet2.Cells["B4"].LoadFromDataTable(dataTable6, true);
                            }
                            #endregion

                            #region No Compliance Overall Report
                            using (SqlConnection sqlConn = new SqlConnection(connectionString))
                            {
                                using (SqlCommand sqlCmd = new SqlCommand("ReportNocompliance", sqlConn))
                                {
                                    sqlCmd.CommandType = CommandType.StoredProcedure;
                                    sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                    sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                    sqlCmd.Parameters.AddWithValue("@machineId", item.MachineId);
                                    sqlConn.Open();
                                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                    {
                                        sqlAdapter.Fill(dataTable3);
                                    }
                                }
                            }

                            dataTable7.Merge(dataTable3);
                            for (int t = 0; t < dataTable7.Rows.Count; t++)
                            {
                                //int l = t;
                                //dataTable4.Rows[0][0] = l + 1;
                                dataTable7.Rows[t][0] = t + 1;
                            }

                            int count4 = dataTable7.Rows.Count;
                            int count3 = dataTable3.Rows.Count;
                            if (count4 > 0)
                            {
                                worksheet3.Cells["B4"].LoadFromDataTable(dataTable7, true);
                            }

                            #endregion

                            #region Graph

                            #region Chart Data

                            int columnNumber = 1;
                            int alpha = 1;
                            worksheet4.Cells["A1"].Value = "Date";
                            worksheet4.Cells["A2"].Value = "NoCode";
                            worksheet4.Cells["A3"].Value = "UnAssignedWO";
                            worksheet4.Cells["A4"].Value = "NoLogin";

                            worksheet4.Cells["A5"].Value = "NoCodePercentage";
                            worksheet4.Cells["A6"].Value = "UnAssignedWOPercentage";
                            worksheet4.Cells["A7"].Value = "NoLoginPercentage";
                            string finalAplha = "";
                            for (int i = 0; i < count3; i++)
                            {
                                //string alphabet = CommonFunction.GetExcelColumnName(i + 2);
                                string alphabet = GetExcelColumnName(i + 2);
                                string date = dataTable3.Rows[i][1].ToString();
                                string noCode = dataTable3.Rows[i][4].ToString();
                                string unAssignedWO = dataTable3.Rows[i][5].ToString();
                                string noLogin = dataTable3.Rows[i][6].ToString();
                                string noCodePercentage = dataTable3.Rows[i][10].ToString();
                                string unAssignedWOPercentage = dataTable3.Rows[i][11].ToString();
                                string noLoginPercentage = dataTable3.Rows[i][12].ToString();

                                #region Excel Graph Data
                                worksheet4.Cells[alphabet + 1].Value = date;
                                worksheet4.Cells[alphabet + 5].Value = date;
                                decimal zeroValue = 0;
                                #region Chart 1


                                if (noCode != "")
                                {
                                    worksheet4.Cells[alphabet + 2].Value = Convert.ToInt32(noCode);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 2].Value = zeroValue;
                                }
                                if (unAssignedWO != "")
                                {
                                    worksheet4.Cells[alphabet + 3].Value = Convert.ToInt32(unAssignedWO);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 3].Value = zeroValue;
                                }
                                if (noLogin != "")
                                {
                                    worksheet4.Cells[alphabet + 4].Value = Convert.ToInt32(noLogin);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 4].Value = zeroValue;
                                }

                                #endregion
                                #region Chart 2

                                if (noCodePercentage != "")
                                {
                                    worksheet4.Cells[alphabet + 5].Value = Convert.ToDecimal(noCodePercentage);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 5].Value = zeroValue;
                                }
                                if (unAssignedWOPercentage != "")
                                {
                                    worksheet4.Cells[alphabet + 6].Value = Convert.ToDecimal(unAssignedWOPercentage);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 6].Value = zeroValue;
                                }
                                if (noLoginPercentage != "")
                                {
                                    worksheet4.Cells[alphabet + 7].Value = Convert.ToDecimal(noLoginPercentage);
                                }
                                else
                                {
                                    worksheet4.Cells[alphabet + 7].Value = zeroValue;
                                }


                                #endregion

                                #endregion
                                //finalAplha = CommonFunction.GetExcelColumnName(alpha + 1);
                                finalAplha = GetExcelColumnName(alpha + 1);
                                alpha++;
                                columnNumber++;
                            }

                            ExcelRange rng = worksheet4.Cells["B1:B" + count3 + 1];
                            rng.Style.Numberformat.Format = "#";
                            //worksheet4.Cells[2, count3 + 2].Style.Numberformat.Format = "##0.0";
                            #endregion

                            #region Chart Graph
                            try
                            {
                                for (int i = 2; i <= 4; i++)
                                {
                                    //string alphabet = CommonFunction.GetExcelColumnName(alpha + 1);
                                    string alphabet = GetExcelColumnName(alpha + 1);
                                    #region chart 1
                                    chart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
                                    chart.Legend.Add();
                                    //chart.SetPosition(SerialNo + 8, 0, 1, 0);
                                    chart.SetPosition(0, 0, 0, 0);
                                    chart.SetSize(count3 * 100, 500);
                                    chart.DataLabel.ShowValue = true;

                                    var series1 = chart.Series.Add($"B{i}:" + finalAplha + $"{i}", "B1:" + finalAplha + "1");
                                    series1.Header = worksheet4.Cells[$"A{i}"].Value.ToString();

                                    chart.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
                                    chart.Title.Text = "No Compliance Value Report";
                                    chart.XAxis.Title.Text = "Day";
                                    chart.YAxis.Title.Text = "Value";
                                    #endregion
                                    #region chart 2
                                    chart1.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
                                    chart1.Legend.Add();
                                    //chart.SetPosition(SerialNo + 8, 0, 1, 0);
                                    chart1.SetPosition(26, 0, 0, 0);
                                    chart1.SetSize(count3 * 100, 500);
                                    chart1.DataLabel.ShowValue = true;

                                    int j = i + 3;
                                    var series2 = chart1.Series.Add($"B{j}:" + finalAplha + $"{j}", "B1:" + finalAplha + "1");

                                    series2.Header = worksheet4.Cells[$"A{j}"].Value.ToString();

                                    chart1.Style = OfficeOpenXml.Drawing.Chart.eChartStyle.Style47;
                                    chart1.Title.Text = "No Compliance Percentage Report";
                                    chart1.XAxis.Title.Text = "Day";
                                    chart1.YAxis.Title.Text = "Percentage";
                                    #endregion

                                }

                            }


                            catch (Exception ex)

                            {

                            }
                            #endregion



                            #endregion

                            #endregion

                            #endregion

                            #region Save and Download

                            p.Save();

                            //Downloding Excel
                            string path = retrivalPath;

                            #endregion

                            obj.isStatus = true;
                            obj.response = path;
                        }
                    }
                    catch (Exception e)
                    {
                        log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                        obj.response = ResourceResponse.ExceptionMessage;
                    }
                }
                catch (Exception e)
                {
                    log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                    //ErrorLog.SendErrorToDB(e);
                    obj.response = ResourceResponse.ExceptionMessage; ;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.response = ResourceResponse.ExceptionMessage;
                obj.isStatus = false;
            }

            return obj;
        }

        public CommonResponse1 OEEDeckFormatReport(OEEDeckFormat data)
        {
            CommonResponse1 obj = new CommonResponse1();

            try
            {
                #region read date time
                DateTime fromDate = DateTime.Now;
                try
                {
                    string[] dt = data.fromDate.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    fromDate = Convert.ToDateTime(frDate);
                }
                catch
                {
                    fromDate = Convert.ToDateTime(data.fromDate);
                }
                DateTime toDate = DateTime.Now;
                try
                {
                    string[] dt = data.toDate.Split('/');
                    string torDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    toDate = Convert.ToDateTime(torDate);
                }
                catch
                {
                    toDate = Convert.ToDateTime(data.toDate).AddHours(24);
                }

                bool isCondition = false;

                if ((data.fromDate != null) && (data.toDate != null))
                {
                    isCondition = true;
                }


                #endregion

                #region Excel and Stuff
                DateTime frda = DateTime.Now;
                string ImageUrlSave = configuration.GetSection("MySettings").GetSection("ImageUrlSave").Value;
                string ImageUrl = configuration.GetSection("MySettings").GetSection("ImageUrl").Value;

                String FileDir = ImageUrlSave + "\\" + "OEE Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                String retrivalPath = ImageUrl + "OEE Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                FileInfo newFile = new FileInfo(FileDir);

                if (newFile.Exists)
                {
                    try
                    {
                        newFile.Delete();  // ensures we create a new workbook
                        newFile = new FileInfo(FileDir);
                    }
                    catch (Exception ex)
                    {
                        //ErrorLog.SendErrorToDB(ex);
                        obj.response = ResourceResponse.ExceptionMessage; ;
                    }
                }

                FileInfo templateFile = new FileInfo("C:\\TataReport\\NewTemplates\\OEEDeckReportTemplate.xlsx");
                ExcelPackage templatep = new ExcelPackage(templateFile);
                ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];

                //Using the File for generation and populating it
                ExcelPackage p = null;
                p = new ExcelPackage(newFile);
                ExcelWorksheet worksheet = null;

                //Creating the WorkSheet for populating
                try
                {
                    worksheet = p.Workbook.Worksheets.Add("OEEDeckFormat", Templatews);
                }
                catch (Exception ex) { }


                if (worksheet == null)
                {
                    worksheet = p.Workbook.Worksheets.Add("OEEDeckFormat", Templatews);

                }

                int sheetcount = p.Workbook.Worksheets.Count;
                p.Workbook.Worksheets.MoveToStart(sheetcount - 1);
                worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                #endregion

                #region 

                var getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.IsNormalWc == 0).ToList();
                if (data.machineId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.MachineId == data.machineId && x.IsNormalWc == 0).ToList();
                }
                else if (data.cellId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId && x.IsNormalWc == 0).ToList();
                }
                else if (data.shopId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId && x.IsNormalWc == 0).ToList();
                }
                else if (data.plantId != 0)
                {
                    getMachineDetails = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId && x.IsNormalWc == 0).ToList();
                }

                #endregion

                try
                {
                    #region Stored procedure calling and to check what is the data to be sent for no date search
                    string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                    //SqlConnection SqlConnection = new SqlConnection(connectionString);
                    #endregion

                    DataTable dataTable = new DataTable();

                    try
                    {
                        #region fill data in excel

                        #region call sp

                        foreach (var item1 in getMachineDetails)
                        {
                            #region Report OEE Deck
                            using (SqlConnection sqlConn = new SqlConnection(connectionString))
                            {
                                
                                using (SqlCommand sqlCmd = new SqlCommand("ReportOEEDeck", sqlConn))
                                {
                                    sqlCmd.CommandType = CommandType.StoredProcedure;
                                    sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                    sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                    sqlCmd.Parameters.AddWithValue("@machineId", item1.MachineId);
                                    sqlConn.Open();
                                    using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                    {
                                        sqlAdapter.Fill(dataTable);
                                    }
                                }
                            }

                            int count = dataTable.Rows.Count;
                            if (count > 0)
                            {
                                List<OEEloadSQLTable> oEEloadSQLTables = new List<OEEloadSQLTable>();
                                for (int i1 = 0; i1 < count; i1++)
                                {
                                    try
                                    {
                                        OEEloadSQLTable oEEloadSQLTable = new OEEloadSQLTable();
                                        string weekStartDateTime = dataTable.Rows[i1][0].ToString();
                                        oEEloadSQLTable.weekStartDateTime = Convert.ToDateTime(weekStartDateTime);
                                        oEEloadSQLTable.machineId = Convert.ToInt32(dataTable.Rows[i1][1].ToString());
                                        oEEloadSQLTable.machineInvNo = Convert.ToString(dataTable.Rows[i1][2].ToString());
                                        oEEloadSQLTable.machineDispName = Convert.ToString(dataTable.Rows[i1][3].ToString());
                                        oEEloadSQLTable.targetOEE = Convert.ToDecimal(dataTable.Rows[i1][4].ToString());
                                        oEEloadSQLTable.durationInMinForWeek = Convert.ToInt32(dataTable.Rows[i1][5].ToString());
                                        oEEloadSQLTable.avaliabilityForWeek = Convert.ToDecimal(dataTable.Rows[i1][6].ToString());
                                        oEEloadSQLTable.weekNumber = Convert.ToInt32(dataTable.Rows[i1][7].ToString());
                                        oEEloadSQLTable.timeTookMinForWeek = Convert.ToDecimal(dataTable.Rows[i1][8].ToString());
                                        oEEloadSQLTable.PerformanceForWeek = Convert.ToDecimal(dataTable.Rows[i1][9].ToString());
                                        oEEloadSQLTable.OEE = Convert.ToDecimal(dataTable.Rows[i1][10].ToString());
                                        oEEloadSQLTables.Add(oEEloadSQLTable);
                                    }
                                    catch (Exception ex)
                                    { }
                                }

                                var weekList = oEEloadSQLTables.Select(m => new { m.weekNumber }).OrderBy(m => m.weekNumber).Distinct();
                                int alpha = 1;
                                foreach (var item in weekList)
                                {
                                    //string alphabet = CommonFunction.GetExcelColumnName(alpha + 4);
                                    string alphabet = GetExcelColumnName(alpha + 4);
                                    worksheet.Cells[alphabet + 1].Value = "Week " + item.weekNumber;
                                    alpha++;
                                }

                                var macDetails = oEEloadSQLTables.Select(m => m.machineId).Distinct();
                                int colCount = 4 + weekList.Count();
                                int i = 2;
                                foreach (var item in macDetails)
                                {
                                    alpha = 1;
                                    var colItems = oEEloadSQLTables.Where(m => m.machineId == item).OrderBy(m => m.weekNumber).ToList();
                                    var TargetOee = oEEloadSQLTables.Where(m => m.machineId == item).Select(m => m.targetOEE).Distinct();
                                    decimal yeildtargetOee = colItems.Select(m => m.OEE).Sum() / colItems.Count();
                                    string ytq = yeildtargetOee.ToString("#.##");
                                    if (ytq != "")
                                    {
                                        //worksheet.Cells["D" + 1].Value = "Yeild Qty";
                                        worksheet.Cells["D" + i].Value = Convert.ToDecimal(ytq);
                                    }
                                    else
                                    {
                                        //worksheet.Cells["D" + 1].Value = "Yeild Qty";
                                        worksheet.Cells["D" + i].Value = 0;
                                    }

                                    if(TargetOee != null)
                                    {
                                        worksheet.Cells["C" + i].Value = TargetOee;
                                    }
                                    else
                                    {
                                        //worksheet.Cells["D" + 1].Value = "Yeild Qty";
                                        worksheet.Cells["C" + i].Value = 0;
                                    }
                                    //worksheet.Cells["A" + 1].Value = "Machine Inv No";
                                    //worksheet.Cells["B" + 1].Value = "Machine Display Name";
                                    foreach (var colItem in colItems)
                                    {
                                        //string alphabet = CommonFunction.GetExcelColumnName(alpha);
                                        string alphabet = GetExcelColumnName(alpha);
                                        if (alpha == 1)
                                        {
                                            worksheet.Cells[alphabet + i].Value = colItem.machineInvNo;
                                            alpha = alpha + 1;
                                            //alphabet = CommonFunction.GetExcelColumnName(alpha);
                                            alphabet = GetExcelColumnName(alpha);
                                            worksheet.Cells[alphabet + i].Value = colItem.machineDispName;
                                            alpha = alpha + 3;
                                            alphabet = GetExcelColumnName(alpha);
                                            worksheet.Cells[alphabet + i].Value = colItem.OEE;
                                            
                                        }
                                        else
                                        {

                                            //alphabet = CommonFunction.GetExcelColumnName(alpha);
                                            alphabet = GetExcelColumnName(alpha);
                                            worksheet.Cells[alphabet + i].Value = colItem.OEE;
                                        }

                                        alpha++;
                                    }
                                    if ((colItems.Count()) < (weekList.Count()))
                                    {
                                        int looper = (weekList.Count()) - (colItems.Count()) + 1;
                                        for (int j = 0; j < looper; j++)
                                        {
                                            //string alphabet = CommonFunction.GetExcelColumnName(alpha);
                                            string alphabet = GetExcelColumnName(alpha);
                                            alpha++;
                                            worksheet.Cells[alphabet + i].Value = 0;
                                        }
                                    }
                                    i++;
                                }

                                //string alphabets = CommonFunction.GetExcelColumnName(colCount);
                                string alphabets = GetExcelColumnName(colCount);
                                Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#34b4eb");
                                worksheet.Cells["A1:" + alphabets + "1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells["A1:" + alphabets + "1"].Style.Fill.BackgroundColor.SetColor(colFromHex);

                                string modelRange = "A1:" + alphabets + (i - 1);
                                var modelTable = worksheet.Cells[modelRange];

                                // Assign borders
                                modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                modelTable.AutoFitColumns();
                            }
                            #endregion

                            #endregion

                            #endregion

                            #region Save and Download

                            p.Save();

                            //Downloding Excel
                            string path = retrivalPath;

                            #endregion

                            obj.isStatus = true;
                            obj.response = path;
                        }
                    }
                    catch (Exception e)
                    {
                        log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                        obj.response = ResourceResponse.ExceptionMessage;
                    }
                }
                catch (Exception e)
                {
                    log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                    //ErrorLog.SendErrorToDB(e);
                    obj.response = ResourceResponse.ExceptionMessage; ;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.response = ResourceResponse.ExceptionMessage;
                obj.isStatus = false;
            }

            return obj;
        }

        public CommonResponse1 OEEDeckFormatLossReasonReport(OEEDeckFormat data)
        {
            CommonResponse1 obj = new CommonResponse1();

            try
            {
                #region read date time
                DateTime fromDate = DateTime.Now;
                try
                {
                    string[] dt = data.fromDate.Split('/');
                    string frDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    fromDate = Convert.ToDateTime(frDate);
                }
                catch
                {
                    fromDate = Convert.ToDateTime(data.fromDate);
                }
                DateTime toDate = DateTime.Now;
                try
                {
                    string[] dt = data.toDate.Split('/');
                    string torDate = dt[2] + '-' + dt[1] + '-' + dt[0];
                    toDate = Convert.ToDateTime(torDate);
                }
                catch
                {
                    toDate = Convert.ToDateTime(data.toDate).AddHours(24);
                }

                bool isCondition = false;

                if ((data.fromDate != null) && (data.toDate != null))
                {
                    isCondition = true;
                }


                #endregion

                #region Excel and Stuff
                DateTime frda = DateTime.Now;
                string ImageUrlSave = configuration.GetSection("AppSettings").GetSection("ImageUrlSave").Value;
                string ImageUrl = configuration.GetSection("AppSettings").GetSection("ImageUrl").Value;

                String FileDir = ImageUrlSave + "\\" + "OEE Loss Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                String retrivalPath = ImageUrl + "OEE Loss Report " + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                FileInfo newFile = new FileInfo(FileDir);

                if (newFile.Exists)
                {
                    try
                    {
                        newFile.Delete();  // ensures we create a new workbook
                        newFile = new FileInfo(FileDir);
                    }
                    catch (Exception ex)
                    {
                        //ErrorLog.SendErrorToDB(ex);
                        obj.response = ResourceResponse.ExceptionMessage; ;
                    }
                }

                FileInfo templateFile = new FileInfo("C:\\TataReport\\NewTemplates\\OEEDeckLossReportTemplate.xlsx");
                ExcelPackage templatep = new ExcelPackage(templateFile);
                ExcelWorksheet Templatews = templatep.Workbook.Worksheets[0];

                //Using the File for generation and populating it
                ExcelPackage p = null;
                p = new ExcelPackage(newFile);
                ExcelWorksheet worksheet = null;

                //Creating the WorkSheet for populating
                try
                {
                    worksheet = p.Workbook.Worksheets.Add("OEEDeckFormatLoss", Templatews);
                }
                catch (Exception ex) { }


                if (worksheet == null)
                {
                    worksheet = p.Workbook.Worksheets.Add("OEEDeckFormatLoss", Templatews);

                }

                int sheetcount = p.Workbook.Worksheets.Count;
                p.Workbook.Worksheets.MoveToStart(sheetcount - 1);
                worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                #endregion

                try
                {
                    #region Stored procedure calling and to check what is the data to be sent for no date search
                    string connectionString = configuration.GetSection("MySettings").GetSection("DbConnection").Value;
                    //SqlConnection SqlConnection = new SqlConnection(connectionString);
                    #endregion

                    DataTable dataTable = new DataTable();

                    try
                    {
                        #region fill data in excel

                        #region call sp

                        #region Report OEE Deck
                        using (SqlConnection sqlConn = new SqlConnection(connectionString))
                        {
                            using (SqlCommand sqlCmd = new SqlCommand("ReportOEEDeckLossReason", sqlConn))
                            {
                                sqlCmd.CommandType = CommandType.StoredProcedure;
                                sqlCmd.Parameters.AddWithValue("@fromDate", fromDate);
                                sqlCmd.Parameters.AddWithValue("@toDate", toDate);
                                sqlCmd.Parameters.AddWithValue("@machineId", data.machineId);
                                sqlConn.Open();
                                using (SqlDataAdapter sqlAdapter = new SqlDataAdapter(sqlCmd))
                                {
                                    sqlAdapter.Fill(dataTable);
                                }
                            }
                        }

                        var machineName = db.Tblmachinedetails.Where(m => m.MachineId == data.machineId).Select(m => m.MachineDispName).FirstOrDefault();

                        worksheet.Cells["C" + 1 + ":" + "BJ" + 1].Merge = true;
                        worksheet.Cells["C" + 1].Value = machineName;

                        int count = dataTable.Rows.Count;
                        if (count > 0)
                        {
                            List<OEELossloadSQLTable> oEEloadSQLTables = new List<OEELossloadSQLTable>();
                            for (int i1 = 0; i1 < count; i1++)
                            {
                                try
                                {
                                    OEELossloadSQLTable oEEloadSQLTable = new OEELossloadSQLTable();
                                    string weekStartDateTime = dataTable.Rows[i1][0].ToString();
                                    oEEloadSQLTable.OEECatID = Convert.ToInt32(weekStartDateTime);
                                    oEEloadSQLTable.OEECatName = Convert.ToString(dataTable.Rows[i1][1].ToString());
                                    oEEloadSQLTable.targetInHrs = Convert.ToInt32(dataTable.Rows[i1][2].ToString());
                                    oEEloadSQLTable.machineID = Convert.ToInt32(dataTable.Rows[i1][3].ToString());
                                    oEEloadSQLTable.weekNumber = Convert.ToInt32(dataTable.Rows[i1][4].ToString());
                                    oEEloadSQLTable.weekStartyear = Convert.ToInt32(dataTable.Rows[i1][5].ToString());
                                    oEEloadSQLTable.lossDurationinMinForWeek = Convert.ToInt32(dataTable.Rows[i1][6].ToString());
                                    oEEloadSQLTable.totalLossDurationinMinForWeek = Convert.ToInt32(dataTable.Rows[i1][7].ToString());
                                    oEEloadSQLTable.overrallLossPOBDMLL = Convert.ToInt32(dataTable.Rows[i1][8].ToString());
                                    oEEloadSQLTable.overallContribution = Convert.ToDecimal(dataTable.Rows[i1][9].ToString());

                                    oEEloadSQLTables.Add(oEEloadSQLTable);
                                }
                                catch (Exception ex)
                                { }
                            }

                            var weekList = oEEloadSQLTables.Select(m => new { m.weekNumber, m.weekStartyear }).OrderBy(m => m.weekStartyear).OrderBy(m => m.weekNumber).Distinct();
                            int alpha = 1;
                            //string alphabet = CommonFunction.GetExcelColumnName(alpha);
                            string alphabet = GetExcelColumnName(alpha);
                            int i11 = 0;
                            foreach (var item in weekList)
                            {
                                try
                                {
                                    //alphabet = CommonFunction.GetExcelColumnName(alpha + 6);
                                    alphabet = GetExcelColumnName(alpha + 6);
                                    string merger = alphabet;
                                    worksheet.Cells[merger + 3].Value = "Hrs";
                                    alpha++;
                                    //alphabet = CommonFunction.GetExcelColumnName(alpha + 6);
                                    alphabet = GetExcelColumnName(alpha + 6);
                                    worksheet.Cells[merger + 2 + ":" + alphabet + 2].Merge = true;
                                    worksheet.Cells[merger + 2].Value = "Week " + item.weekNumber + " " + item.weekStartyear;
                                    worksheet.Cells[alphabet + 3].Value = "% of Contrn";
                                    alpha++;
                                }
                                catch (Exception ex)
                                {

                                }
                                i11++;
                            }
                            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#34b4eb");
                            worksheet.Cells["A1:" + alphabet + "3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A1:" + alphabet + "3"].Style.Fill.BackgroundColor.SetColor(colFromHex);
                            var categoryDetails = oEEloadSQLTables.Select(m => m.OEECatName).Distinct();
                            string apf = alphabet;
                            int i2 = 4;

                            int loop = 1;
                            foreach (var item in categoryDetails)
                            {
                                alpha = 1;
                                worksheet.Cells["A" + i2].Value = loop;
                                worksheet.Cells["B" + i2].Value = item;

                                var catItems = oEEloadSQLTables.Where(m => m.OEECatName == item).OrderBy(m => m.weekStartyear).OrderBy(m => m.weekNumber).ToList();
                                decimal hrs = 0;
                                decimal contrn = 0;
                                foreach (var catItem in catItems)
                                {
                                    worksheet.Cells["c" + i2].Value = catItem.targetInHrs;
                                    //alphabet = CommonFunction.GetExcelColumnName(alpha + 6);
                                    alphabet = GetExcelColumnName(alpha + 6);
                                    string merger = alphabet;
                                    decimal lossDurationinMinForWeek = Convert.ToDecimal(catItem.lossDurationinMinForWeek / 60.00);
                                    hrs = hrs + lossDurationinMinForWeek;
                                    worksheet.Cells[merger + i2].Value = Convert.ToDecimal(lossDurationinMinForWeek.ToString("#.##0"));
                                    alpha++;
                                    //alphabet = CommonFunction.GetExcelColumnName(alpha + 6);
                                    alphabet = GetExcelColumnName(alpha + 6);
                                    decimal overallContribution = catItem.overallContribution;
                                    contrn = contrn + overallContribution;
                                    worksheet.Cells[alphabet + i2].Value = Convert.ToDecimal(overallContribution.ToString("#.##0"));
                                    alpha++;
                                }
                                if ((catItems.Count() * 2) < (weekList.Count() * 2))
                                {
                                    int looper = (weekList.Count() * 2) - (catItems.Count() * 2);
                                    for (int i = 0; i < looper; i++)
                                    {
                                        //alphabet = CommonFunction.GetExcelColumnName(alpha + 6);
                                        alphabet = GetExcelColumnName(alpha + 6);
                                        alpha++;
                                        worksheet.Cells[alphabet + i2].Value = 0;
                                    }
                                }
                                worksheet.Cells["E" + i2].Value = hrs;
                                worksheet.Cells["F" + i2].Value = contrn;
                                loop++;
                                i2++;
                            }


                            string modelRange = "A1:" + apf + (loop + 2);
                            var modelTable = worksheet.Cells[modelRange];

                            // Assign borders
                            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                            // Fill worksheet with data to export
                            modelTable.AutoFitColumns();

                        }
                        #endregion

                        #endregion

                        #endregion

                        #region Save and Download

                        p.Save();

                        //Downloding Excel
                        string path = retrivalPath;

                        #endregion

                        obj.isStatus = true;
                        obj.response = path;
                    }
                    catch (Exception e)
                    {
                        log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                        obj.response = ResourceResponse.ExceptionMessage;
                    }
                }
                catch (Exception e)
                {
                    log.Error(e); if (e.InnerException != null) { log.Error(e.InnerException.ToString()); }
                    //ErrorLog.SendErrorToDB(e);
                    obj.response = ResourceResponse.ExceptionMessage; ;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.response = ResourceResponse.ExceptionMessage;
                obj.isStatus = false;
            }

            return obj;
        }

        public CommonResponse UpdateLoginDetails(long machineId)
        {
            CommonResponse commonResponse = new CommonResponse();

            try
            {
                #region Get Corrected Date
                string correctedDate = null;
                var StartTime = db.Tbldaytiming.Where(m => m.IsDeleted == 0).FirstOrDefault();
                TimeSpan End = StartTime.EndTime;
                TimeSpan EndTimeSpan = new TimeSpan(0, 0, 0);
                TimeSpan TimeSpanNow = DateTime.Now.TimeOfDay;
                if (TimeSpanNow >= EndTimeSpan && TimeSpanNow <= End)
                {
                    correctedDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                }
                else
                {
                    correctedDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                #endregion

                #region Get Shift
                DateTime Time = DateTime.Now;
                TimeSpan Tm = new TimeSpan(Time.Hour, Time.Minute, Time.Second);

                //CommonFunction cf = new CommonFunction();
                List<ShiftList> TblshiftMstr = ShiftList(correctedDate);

                var shiftDetaild = TblshiftMstr.Where(m => m.shiftStartDateTime <= Time && m.shiftEndDateTime >= Time).FirstOrDefault();
                var shiftDetails = db.TblshiftMstr.Where(m => m.ShiftId == shiftDetaild.shfitId).FirstOrDefault();

                string shift = "";

                if (shiftDetails != null)
                {
                    shift = shiftDetails.ShiftName;
                }

                #endregion


                #region Logic to add and update


                var loginDetails = db.LoginDetails.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.Shift == shift && m.IsCompleted == false).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();
                if (loginDetails == null)
                {
                    #region Algorithm to end Shifts based on machine and shift and corrected date
                    string shftstarttime = correctedDate + " " + shiftDetails.StartTime;
                    DateTime shftStDt = Convert.ToDateTime(shftstarttime);
                    var loginDetaild = db.LoginDetails.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.Shift == shift && m.IsCompleted == false && m.StartTime == shftStDt && m.EndTime == shftStDt).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();
                    if (loginDetaild == null)
                    {

                        var dbCheck = db.LoginDetails.Where(m => m.MachineId == machineId && m.IsCompleted == false).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();
                        if (dbCheck != null)
                        {
                            dbCheck.IsCompleted = true;
                            db.SaveChanges();
                        }
                        if (shiftDetaild.shfitId == 1)
                        {
                            // do nothing
                        }
                        else if (shiftDetaild.shfitId == 2)
                        {
                            #region insert shift A
                            string shftstarttime1 = correctedDate + " 06:00:00";
                            DateTime shftStDt1 = Convert.ToDateTime(shftstarttime1);
                            var loginDetaildA = db.LoginDetails.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.Shift == "A" && m.StartTime == shftStDt1 && m.EndTime == shftStDt1).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();

                            if (loginDetaildA == null)
                            {
                                #region add login details
                                LoginDetails loginDetais11 = new LoginDetails();
                                loginDetais11.Shift = "A";
                                loginDetais11.CorrectedDate = correctedDate;
                                loginDetais11.StartTime = shftStDt1;
                                loginDetais11.EndTime = shftStDt1;
                                loginDetais11.CreatedOn = DateTime.Now;
                                loginDetais11.IsDeleted = false;
                                loginDetais11.IsActive = true;
                                loginDetais11.IsCompleted = true;
                                loginDetais11.MachineId = machineId;
                                db.LoginDetails.Add(loginDetais11);
                                db.SaveChanges();
                                #endregion
                            }
                            #endregion
                        }
                        else if (shiftDetaild.shfitId == 3)
                        {

                            #region insert shift A
                            string shftstarttime1 = correctedDate + " 06:00:00";
                            DateTime shftStDt1 = Convert.ToDateTime(shftstarttime1);
                            var loginDetaildA = db.LoginDetails.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.Shift == "A" && m.StartTime == shftStDt1 && m.EndTime == shftStDt1).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();

                            if (loginDetaildA == null)
                            {
                                #region add login details
                                LoginDetails loginDetais11 = new LoginDetails();
                                loginDetais11.Shift = "A";
                                loginDetais11.CorrectedDate = correctedDate;
                                loginDetais11.StartTime = shftStDt1;
                                loginDetais11.EndTime = shftStDt1;
                                loginDetais11.CreatedOn = DateTime.Now;
                                loginDetais11.IsDeleted = false;
                                loginDetais11.IsActive = true;
                                loginDetais11.IsCompleted = true;
                                loginDetais11.MachineId = machineId;
                                db.LoginDetails.Add(loginDetais11);
                                db.SaveChanges();
                                #endregion
                            }
                            #endregion

                            #region insert shift B
                            string shftstarttime12 = correctedDate + " 14:00:00";
                            DateTime shftStDt12 = Convert.ToDateTime(shftstarttime12);
                            var loginDetaildB = db.LoginDetails.Where(m => m.MachineId == machineId && m.CorrectedDate == correctedDate && m.Shift == "B" && m.StartTime == shftStDt12 && m.EndTime == shftStDt12).OrderByDescending(m => m.LoginDetailsId).FirstOrDefault();

                            if (loginDetaildB == null)
                            {
                                #region add login details
                                LoginDetails loginDetais112 = new LoginDetails();
                                loginDetais112.Shift = "B";
                                loginDetais112.CorrectedDate = correctedDate;
                                loginDetais112.StartTime = shftStDt12;
                                loginDetais112.EndTime = shftStDt12;
                                loginDetais112.CreatedOn = DateTime.Now;
                                loginDetais112.IsDeleted = false;
                                loginDetais112.IsActive = true;
                                loginDetais112.IsCompleted = true;
                                loginDetais112.MachineId = machineId;
                                db.LoginDetails.Add(loginDetais112);
                                db.SaveChanges();
                                #endregion
                            }
                            #endregion
                        }
                        #region insert empty row at the begining of every shift
                        #region add login details
                        LoginDetails loginDetaisl = new LoginDetails();
                        loginDetaisl.Shift = shift;
                        loginDetaisl.CorrectedDate = correctedDate;
                        loginDetaisl.StartTime = shftStDt;
                        loginDetaisl.EndTime = shftStDt;
                        loginDetaisl.CreatedOn = DateTime.Now;
                        loginDetaisl.IsDeleted = false;
                        loginDetaisl.IsActive = true;
                        loginDetaisl.IsCompleted = true;
                        loginDetaisl.MachineId = machineId;
                        db.LoginDetails.Add(loginDetaisl);
                        db.SaveChanges();
                        #endregion
                        #endregion
                    }
                    #endregion

                    #region add login details
                    LoginDetails loginDetail = new LoginDetails();
                    loginDetail.Shift = shift;
                    loginDetail.CorrectedDate = correctedDate;
                    loginDetail.StartTime = DateTime.Now;
                    loginDetail.EndTime = DateTime.Now;
                    loginDetail.CreatedOn = DateTime.Now;
                    loginDetail.IsDeleted = false;
                    loginDetail.IsActive = true;
                    loginDetail.IsCompleted = false;
                    loginDetail.MachineId = machineId;
                    db.LoginDetails.Add(loginDetail);
                    db.SaveChanges();
                    #endregion
                }
                else
                {
                    #region Update login details

                    DateTime lastLoginDateTime = Convert.ToDateTime(loginDetails.EndTime);
                    double diffrenceSecond = (DateTime.Now - lastLoginDateTime).TotalSeconds;
                    if (diffrenceSecond >= 12)
                    {
                        #region Update last row end time to current date and add new row with start date with current start date
                        #region Update
                        loginDetails.IsCompleted = true;
                        loginDetails.ModifiedOn = DateTime.Now;
                        db.SaveChanges();
                        #endregion
                        #region Add
                        LoginDetails loginDetail = new LoginDetails();
                        loginDetail.Shift = shift;
                        loginDetail.CorrectedDate = correctedDate;
                        loginDetail.StartTime = DateTime.Now;
                        loginDetail.EndTime = DateTime.Now;
                        loginDetail.CreatedOn = DateTime.Now;
                        loginDetail.IsDeleted = false;
                        loginDetail.IsActive = true;
                        loginDetail.IsCompleted = false;
                        loginDetail.MachineId = machineId;
                        db.LoginDetails.Add(loginDetail);
                        db.SaveChanges();
                        #endregion
                        #endregion

                    }
                    else
                    {
                        #region Update last rows end time to current date
                        loginDetails.IsCompleted = false;
                        loginDetails.EndTime = DateTime.Now;
                        db.SaveChanges();
                        #endregion
                    }

                    #endregion

                }



                #endregion

                commonResponse.isTure = true;
                commonResponse.response = "Updated Successfully";
            }
            catch (Exception ex)
            {
                commonResponse.isTure = false;
                commonResponse.response = "Exception while updating";
            }

            return commonResponse;

        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public string NumberSuffix(int num)
        {
            int ones = num % 10;
            int tens = (num / 10) % 10;
            string suff = "";
            if (tens == 1)
            {
                suff = "th";
            }
            else
            {
                switch (ones)
                {
                    case 1: suff = "st"; break;
                    case 2: suff = "nd"; break;
                    case 3: suff = "rd"; break;
                    default: suff = "th"; break;
                }
            }
            return num + suff;
        }

        public List<ShiftList> ShiftList(string correctedDate)
        {
            List<ShiftList> shiftLists = new List<ShiftList>();
            try
            {
                var shiftDetails = db.TblshiftMstr.ToList();
                int i = 1;
                foreach (var item in shiftDetails)
                {
                    ShiftList shiftList = new ShiftList();
                    shiftList.shfitId = item.ShiftId;
                    shiftList.shiftName = item.ShiftName;

                    switch (i)
                    {
                        case 1:
                            string datee = DateTime.Now.Date.ToShortDateString();
                            string date1 = correctedDate + " " + item.StartTime;
                            string dateee = Convert.ToDateTime(correctedDate).ToShortDateString();
                            string date2 = dateee + " " + item.EndTime;
                            shiftList.shiftStartDateTime = Convert.ToDateTime(date1);
                            shiftList.shiftEndDateTime = Convert.ToDateTime(date2);
                            break;
                        case 2:
                            string datee1 = DateTime.Now.Date.ToShortDateString();
                            string date11 = correctedDate + " " + item.StartTime;
                            string dateee1 = Convert.ToDateTime(correctedDate).ToShortDateString();
                            string date21 = dateee1 + " " + item.EndTime;
                            shiftList.shiftStartDateTime = Convert.ToDateTime(date11);
                            shiftList.shiftEndDateTime = Convert.ToDateTime(date21);
                            break;
                        case 3:
                            string datee2 = DateTime.Now.Date.ToShortDateString();
                            string date12 = correctedDate + " " + item.StartTime;
                            string dateee2 = Convert.ToDateTime(correctedDate).Date.AddDays(1).ToShortDateString();
                            string date22 = dateee2 + " " + item.EndTime;
                            shiftList.shiftStartDateTime = Convert.ToDateTime(date12);
                            shiftList.shiftEndDateTime = Convert.ToDateTime(date22);
                            break;
                        default: break;
                    }

                    shiftLists.Add(shiftList);
                    i++;
                }
            }
            catch (Exception ex)
            {
            }

            return shiftLists;
        }

        #endregion

    }
}
