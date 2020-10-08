﻿using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Data.SqlClient;

namespace AutoReportMail
{
    public partial class Form1 : Form
    {
        i_facility_talEntities1 db = new i_facility_talEntities1();
        public Form1()
        {
            InitializeComponent();
            this.Visible = false;

            try
            {
                MyTimer_Tick();
            }
            catch (Exception exception)
            {
                //MessageBox.Show(exception.ToString());
                IntoFile("Main :" + exception);
            }
            //Process[] GetMyProc = Process.GetProcessesByName("AutoReportMail");
            //if (GetMyProc.Count() > 0)
            //{
            //    try
            //    {
            //        //MessageBox.Show("ablut to close");
            //        GetMyProc[0].CloseMainWindow();
            //        System.Threading.Thread.Sleep(10000);
            //        GetMyProc[0].Kill();
            //    }
            //    catch (Exception eclose)
            //    {
            //        //MessageBox.Show("ablut to close");
            //        IntoFile(" " + eclose);
            //        System.Threading.Thread.Sleep(10000);
            //        GetMyProc[0].Kill();
            //    }
            //}
        }

        private void MyTimer_Tick()
        {
            try
            {
                CallMethod();
            }
            catch (Exception ex)
            {
                IntoFile("" + ex);
            }
        }


        public async void CallMethod()
        {
            using (i_facility_talEntities1 dbautorep = new i_facility_talEntities1())
            {
                String StartDate = null;
                String EndDate = null;

                //For Client
                DateTime PresentDate = System.DateTime.Now;
                String PreviousDate = System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");

                ////For Testing
                //DateTime PresentDate = new DateTime(2020, 08, 17, 00, 00, 00);
                //String PreviousDate = PresentDate.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");

                var AutoRepSetList = dbautorep.tbl_autoreportsetting.Where(m => m.IsDeleted == 0).ToList();
                //MessageBox.Show(""+AutoRepSetList.Count);
                foreach (var AutoRow in AutoRepSetList)
                {
                    try
                    {
                        int AutoReportID = (int)AutoRow.AutoReportID;
                        int ReportID = (int)AutoRow.ReportID;
                        IntoFile("ReportId :" + ReportID);
                        int basedon = (int)AutoRow.BasedOn;
                        int AutoTime = (int)AutoRow.AutoReportTimeID;
                        String PlantID = AutoRow.PlantID.ToString();
                        String ShopID = AutoRow.ShopID.ToString();
                        String CellID = AutoRow.CellID.ToString();
                        String WCID = AutoRow.MachineID.ToString();
                        DateTime AutoReportRunTime = (DateTime)AutoRow.NextRunDate;
                        int sendreportID = -1;
                        int SendMail = -1;
                        string TabularType = null;
                        if (AutoRow.AutoReportTimeID == 1)
                        {
                            TabularType = "Day";
                            PreviousDate = System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                        }
                        else if (AutoRow.AutoReportTimeID == 2)
                        {
                            TabularType = "Week";
                            DateTime begining, end;
                            GetWeek(DateTime.Now, new CultureInfo("fr-FR"), out begining, out end);
                            PresentDate = begining.AddDays(-1);
                            PreviousDate = PresentDate.AddDays(-6).ToString("yyyy-MM-dd 00:00:00");
                        }
                        else if (AutoRow.AutoReportTimeID == 3)
                        {
                            TabularType = "Month";
                            DateTime _1stOfCurrentMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 01, 00, 00, 01);
                            PresentDate = _1stOfCurrentMonth.AddDays(-1);
                            PreviousDate = new DateTime(PresentDate.Year, PresentDate.Month, 01, 00, 00, 01).ToString();
                        }
                        else if (AutoRow.AutoReportTimeID == 4)
                        {
                            TabularType = "Year";
                            DateTime _1stOfCurrentYear = new DateTime(DateTime.Now.Year, 01, 01, 00, 00, 01);
                            PresentDate = _1stOfCurrentYear.AddDays(-1);
                            PreviousDate = new DateTime(PresentDate.Year, 01, 01, 00, 00, 01).ToString();
                        }
                        // MessageBox.Show("PresentDate: " + PresentDate);
                        // MessageBox.Show("AutoReportRunTime: " + AutoReportRunTime);
                        // Put While Loop and You can get 
                        if (PresentDate > AutoReportRunTime)
                        {
                            var Logdata = dbautorep.tbl_autoreport_log.Where(n => n.AutoReportID == AutoReportID).OrderByDescending(n => n.AutoReportLogID).FirstOrDefault();
                            if (Logdata != null)
                            {
                                DateTime LogDateTime = (DateTime)Logdata.CorrectedDate;
                                int ExcelCreated = (int)Logdata.ExcelCreated;
                                int MailSent = (int)Logdata.MailSent;
                                if (LogDateTime.Date < PresentDate.Date)
                                {
                                    sendreportID = ReportID;
                                }
                                else if (ExcelCreated == 1 && MailSent == 0)
                                {
                                    SendMail = ReportID;
                                }
                            }
                            else
                            {
                                sendreportID = ReportID;
                            }

                            Task<string> path = null;
                            switch (sendreportID)
                            {
                                case 2:
                                    path = await Task<string>.FromResult(OEEReportExcel(PreviousDate, PresentDate.ToString(), "GH", "OverAll", PlantID, ShopID, CellID, WCID, TabularType));
                                    break;
                                case 3:
                                    path = await Task<string>.FromResult(OEEReportExcel(PreviousDate, PresentDate.ToString(), "NoBlue", "OverAll", PlantID, ShopID, CellID, WCID));
                                    break;
                                case 4:
                                    path = await Task<string>.FromResult(JOBReportExcel(PreviousDate, PresentDate.ToString(), "OverAll", PlantID, ShopID, CellID, WCID, " "));
                                    break;
                                case 5:
                                    path = await Task<string>.FromResult(UtilizationReportExcel(PreviousDate, PresentDate.ToString(), PlantID, ShopID, CellID, WCID, TabularType));
                                    break;
                                case 6:
                                    path = await Task<string>.FromResult(LossAnalysisReportExcel(PreviousDate, PresentDate.ToString(), "OverAll", PlantID, ShopID, CellID, WCID));
                                    break;
                                case 7:
                                    path = await Task<string>.FromResult(MRRReportExcel(PreviousDate, PresentDate.ToString(), PlantID, null, ShopID, CellID, WCID));
                                    break;
                                case 8:
                                    path = await Task<string>.FromResult(ManualWOReportExcel(PreviousDate, PresentDate.ToString(), PlantID, ShopID, CellID, WCID));
                                    break;
                                case 9:
                                    path = await Task<string>.FromResult(NoLoginReportExcel(PreviousDate, PresentDate.ToString(), PlantID, ShopID, CellID, WCID));
                                    break;
                                case 10:
                                    path = await Task<string>.FromResult(UnAssignedWOReportExcel(PreviousDate, PresentDate.ToString(), PlantID, ShopID, CellID, WCID));
                                    break;
                                case 11:
                                    path = await Task<string>.FromResult(UnIdentifiedReportExcel(PreviousDate, PresentDate.ToString(), PlantID, ShopID, CellID, WCID));
                                    break;
                            }
                            if (path != null)
                            {
                                string filepath = Convert.ToString(path.Result);
                                Task<bool> mailSendStatus = sendMail(AutoReportID, DateTime.Now, filepath);
                                DateTime NextDate = DateTime.Now.AddDays(1);
                                //update NextRunDate
                                if (Convert.ToBoolean(mailSendStatus.Result))
                                {
                                    if (AutoRow.AutoReportTimeID == 1) //day
                                    {
                                        AutoRow.NextRunDate = Convert.ToDateTime(DateTime.Now.AddDays(1).ToString("yyyy-MM-dd 08:00:00"));
                                    }
                                    else if (AutoRow.AutoReportTimeID == 2) //week
                                    {
                                        DateTime begining, end;
                                        GetWeek(DateTime.Now, new CultureInfo("fr-FR"), out begining, out end);
                                        AutoRow.NextRunDate = Convert.ToDateTime(end.AddDays(1).ToString("yyyy-MM-dd 08:00:00"));
                                    }
                                    else if (AutoRow.AutoReportTimeID == 3) //month
                                    {
                                        DateTime Temp2 = new DateTime(DateTime.Now.Year, DateTime.Now.AddMonths(1).Month, 01, 08, 00, 00);
                                        AutoRow.NextRunDate = Temp2;
                                    }
                                    else if (AutoRow.AutoReportTimeID == 4) //year
                                    {
                                        DateTime Temp2 = new DateTime(DateTime.Now.AddYears(1).Year, 01, 01, 08, 00, 00);
                                        AutoRow.NextRunDate = Temp2;
                                    }

                                    dbautorep.Entry(AutoRow).State = System.Data.Entity.EntityState.Modified;
                                    dbautorep.SaveChanges();

                                    //MsqlConnection mcUpdateRow = new MsqlConnection();
                                    //mcUpdateRow.open();
                                    //MySqlCommand cmdDeleteRows = new MySqlCommand("UPDATE i_facility_tal.dbo.tbl_autoreportsetting set NextRunDate = '" + NextDate + "' WHERE AutoReportID = '" + AutoReportID + "';", mcUpdateRow.msqlConnection);
                                    //cmdDeleteRows.ExecuteNonQuery();
                                    //mcUpdateRow.close();

                                    tbl_autoreport_log arl = new tbl_autoreport_log();
                                    arl.AutoReportID = AutoReportID;
                                    arl.CompletedOn = DateTime.Now;
                                    arl.CorrectedDate = DateTime.Now.Date;
                                    arl.ExcelCreated = 1;
                                    arl.ExcelCreatedTime = DateTime.Now;
                                    arl.InsertedOn = DateTime.Now;
                                    arl.MailSent = 1;
                                    db.tbl_autoreport_log.Add(arl);
                                    db.SaveChanges();
                                }
                                else
                                {
                                    IntoFile("Mail Send Failed.");
                                }
                            }
                            else
                            {
                                IntoFile("Path is Null. ");
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        IntoFile("One of the AutoReports Failed. " + e);
                    }
                }
            }
        }


        //Report Methods
        public async Task<string> OEEReportExcel(string StartDate, string EndDate, string TimeFactor, string ProdFAI, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null, string TabularType = "Day")
        {
            try
            {
                DateTime frda = DateTime.Now;
                if (string.IsNullOrEmpty(StartDate) == true)
                {
                    StartDate = DateTime.Now.Date.ToString();
                }
                if (string.IsNullOrEmpty(EndDate) == true)
                {
                    EndDate = StartDate;
                }

                DateTime frmDate = Convert.ToDateTime(StartDate);
                DateTime toDate = Convert.ToDateTime(EndDate);

                double TotalDay = toDate.Subtract(frmDate).TotalDays;


                #region MacCount & LowestLevel
                string lowestLevel = null;
                int MacCount = 0;
                int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
                string Header = null;
                if (string.IsNullOrEmpty(WorkCenterID))
                {
                    if (string.IsNullOrEmpty(CellID))
                    {
                        if (string.IsNullOrEmpty(ShopID))
                        {
                            if (string.IsNullOrEmpty(PlantID))
                            {
                                //donothing
                            }
                            else
                            {
                                lowestLevel = "Plant";
                                plantId = Convert.ToInt32(PlantID);
                                MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId && m.IsNormalWC == 0).ToList().Count();
                                var plantName = (from plant in db.tblplants
                                                 where plant.PlantID == plantId
                                                 select new { plantname = plant.PlantName }).SingleOrDefault();

                                Header = plantName.plantname.ToString();
                            }
                        }
                        else
                        {
                            lowestLevel = "Shop";
                            shopId = Convert.ToInt32(ShopID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId && m.IsNormalWC == 0).ToList().Count();
                            var shopName = (from shop in db.tblshops
                                            where shop.ShopID == shopId
                                            select new { shopname = shop.ShopName }).SingleOrDefault();

                            Header = shopName.shopname;
                        }
                    }
                    else
                    {
                        lowestLevel = "Cell";
                        cellId = Convert.ToInt32(CellID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId && m.IsNormalWC == 0).ToList().Count();
                        var cellName = (from cell in db.tblcells
                                        where cell.CellID == cellId
                                        select new { wcname = cell.CellName }).SingleOrDefault();

                        Header = cellName.wcname;
                    }
                }
                else
                {
                    lowestLevel = "WorkCentre";
                    wcId = Convert.ToInt32(WorkCenterID);
                    MacCount = 1;
                    var WCName = (from wc in db.tblmachinedetails
                                  where wc.MachineID == wcId
                                  select new { wcname = wc.MachineDispName }).SingleOrDefault();
                    Header = WCName.wcname;
                }

                #endregion

                FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\OEEReportGodHours.xlsx");
                if (TimeFactor == "GH")
                {
                    templateFile = new FileInfo(@"C:\TataReport\NewTemplates\OEEReportGodHours.xlsx");
                }
                else if (TimeFactor == "NoBlue")
                {
                    templateFile = new FileInfo(@"C:\TataReport\NewTemplates\OEEReportAdjusted.xlsx");
                }

                ExcelPackage templatep = new ExcelPackage(templateFile);
                ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
                ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

                String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
                bool exists = System.IO.Directory.Exists(FileDir);
                if (!exists)
                    System.IO.Directory.CreateDirectory(FileDir);

                FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
                if (TimeFactor == "GH")
                {
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"));
                }
                else if (TimeFactor == "NoBlue")
                {
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportAdjusted" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"));
                }
                if (newFile.Exists)
                {
                    try
                    {
                        newFile.Delete();  // ensures we create a new workbook
                        newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"));
                        if (TimeFactor == "GH")
                        {
                            newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"));
                        }
                        else if (TimeFactor == "NoBlue")
                        {
                            newFile = new FileInfo(System.IO.Path.Combine(FileDir, "OEEReportAdjusted" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"));
                        }
                    }
                    catch
                    {
                        //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                        //return View();
                    }
                }
                //Using the File for generation and populating it
                ExcelPackage p = null;
                p = new ExcelPackage(newFile);
                ExcelWorksheet worksheet = null;
                ExcelWorksheet worksheetGraph = null;

                //Creating the WorkSheet for populating
                try
                {
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                    worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                }
                catch { }

                if (worksheet == null)
                {
                    try{
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                    }
                    catch (Exception e)
                    { }
                    //p.Workbook.Worksheets.Delete("Graphs");
                } 
                if (worksheetGraph == null)
                {
                    try{
                    worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                    }
                    catch (Exception e)
                    { }
                }
                int sheetcount = p.Workbook.Worksheets.Count;
                p.Workbook.Worksheets.MoveToStart(sheetcount);
                //worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                

                #region Get Machines List
                DataTable machin = new DataTable();
                DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = frmDate.ToString("yyyy-MM-dd");
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query1 = null;
                    if (lowestLevel == "Plant")
                    {
                        query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and IsNormalWC = 0 and" +
                            " ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or " +
                            "  ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                        //query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and IsNormalWC = 0 and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                    }
                    else if (lowestLevel == "Shop")
                    {
                        //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                        query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                    }
                    else if (lowestLevel == "Cell")
                    {
                        //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                        query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                    }
                    else if (lowestLevel == "WorkCentre")
                    {
                        //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                        query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                    }

                    //IntoFile(query1);
                    //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(machin);
                    mc.close();
                }
                #endregion

                //DataTable for Consolidated Data 
                DataTable DTConsolidatedLosses = new DataTable();
                DTConsolidatedLosses.Columns.Add("Plant", typeof(string));
                DTConsolidatedLosses.Columns.Add("Shop", typeof(string));
                DTConsolidatedLosses.Columns.Add("Cell", typeof(string));
                DTConsolidatedLosses.Columns.Add("WCInvNo", typeof(string));
                DTConsolidatedLosses.Columns.Add("WCName", typeof(string));
                DTConsolidatedLosses.Columns.Add("CorrectedDate", typeof(string));

                //Add Other Cols of Excel into DataTable
                DTConsolidatedLosses.Columns.Add("OpTime", typeof(double));
                DTConsolidatedLosses.Columns["OpTime"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("AvailableTime", typeof(double));
                DTConsolidatedLosses.Columns["AvailableTime"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("SCTvsPP", typeof(double));
                DTConsolidatedLosses.Columns["SCTvsPP"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("ScrapQtyTime", typeof(double));
                DTConsolidatedLosses.Columns["ScrapQtyTime"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("ReworkTime", typeof(double));
                DTConsolidatedLosses.Columns["ReworkTime"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("CuttingTime", typeof(double));
                DTConsolidatedLosses.Columns["CuttingTime"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("DaysWorking", typeof(double));
                DTConsolidatedLosses.Columns["DaysWorking"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("GodHours", typeof(double));
                DTConsolidatedLosses.Columns["GodHours"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("TotalSTDHours", typeof(double));
                DTConsolidatedLosses.Columns["TotalSTDHours"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("RejectionHours", typeof(double));
                DTConsolidatedLosses.Columns["RejectionHours"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("MinorLoss", typeof(double));
                DTConsolidatedLosses.Columns["MinorLoss"].DefaultValue = 0.0;
                DTConsolidatedLosses.Columns.Add("Breakdown", typeof(double));
                DTConsolidatedLosses.Columns["Breakdown"].DefaultValue = 0.0;

                //Add Cols for A,P,Q and OEE for Individual Dates.
                DTConsolidatedLosses.Columns.Add("Avail", typeof(string));
                DTConsolidatedLosses.Columns.Add("Perf", typeof(string));
                DTConsolidatedLosses.Columns.Add("Qual", typeof(string));
                DTConsolidatedLosses.Columns.Add("OEE", typeof(string));

                //Get All Losses and Insert into DataTable

                DataTable LossCodesData = new DataTable();
                using (MsqlConnection mcLossCodes = new MsqlConnection())
                {
                    mcLossCodes.open();
                    startDateTime = frmDate.ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);
                    //string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where MessageType != 'BREAKDOWN' and ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                    //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "') end) ) and LossCodeID NOT IN (  "
                    //            + "SELECT DISTINCT LossCodeID FROM (  "
                    //            + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where  MessageType != 'BREAKDOWN' and  LossCodesLevel1ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                    //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "' ) end) ) "
                    //            + "UNION  "
                    //            + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where  MessageType != 'BREAKDOWN' and  LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                    //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "' ) end) )  "
                    //            + ") AS derived ) order by LossCodesLevel1ID  ;";


                    string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where MessageType != 'BREAKDOWN' and ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                                + "((IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "') ) ) and LossCodeID NOT IN (  "
                                + "SELECT DISTINCT LossCodeID FROM (  "
                                + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where  MessageType != 'BREAKDOWN' and  LossCodesLevel1ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                                + "( (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "' ) ) ) "
                                + "UNION  "
                                + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where  MessageType != 'BREAKDOWN' and  LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                                + "((IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime + "' ) ) )  "
                                + ") AS derived ) order by LossCodesLevel1ID  ;";


                    //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                    SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                    daLossCodesData.Fill(LossCodesData);
                    mcLossCodes.close();
                }

                //int LossesStartsATCol = 21;
                int LossesStartsATCol = 23;
                var LossesList = new List<KeyValuePair<int, string>>();

                #region LossCodes Into LossList
                for (int i = 0; i < LossCodesData.Rows.Count; i++)
                {
                    int losscode = Convert.ToInt32(LossCodesData.Rows[i][0]);
                    string losscodeName = Convert.ToString(LossCodesData.Rows[i][1]);

                    var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == losscode).FirstOrDefault();
                    int level = lossdata.LossCodesLevel;
                    if (level == 3)
                    {
                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                        int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                        var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                        var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                        losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                    }

                    else if (level == 2)
                    {
                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                        var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                        losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                    }
                    else if (level == 1)
                    {
                        if (losscode == 999)
                        {
                            losscodeName = "NoCode Entered";
                        }
                        else if (losscode == 9999)
                        {
                            losscodeName = "UnIdentified BreakDown";
                        }
                        else
                        {
                            losscodeName = lossdata.LossCode;
                        }
                    }
                    //losscodeName = LossHierarchy3rdLevel(losscode);
                    DTConsolidatedLosses.Columns.Add(losscodeName, typeof(double));
                    DTConsolidatedLosses.Columns[losscodeName].DefaultValue = "0";

                    //Code to write LossesNames to Excel.
                    string columnAlphabet = ExcelColumnFromNumber(LossesStartsATCol);

                    worksheet.Cells[columnAlphabet + 3].Value = losscodeName;
                    worksheet.Cells[columnAlphabet + 4].Value = "AF";

                    LossesStartsATCol++;
                    //Add the LossesToList
                    LossesList.Add(new KeyValuePair<int, string>(losscode, losscodeName));
                }

                #endregion

                DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
                //For each Date ...... for all Machines.
                var Col = 'B';
                int Row = 5 + machin.Rows.Count + 3; // Gap to Insert OverAll data. DataStartRow + MachinesCount + 2(1 for HighestLevel & another for Gap).
                int Sno = 1;
                string finalLossCol = null;

                //for (int i = 0; i < TotalDay + 1; i++)
                int l = 0;
                do
                {
                    DateTime begining = UsedDateForExcel, end = UsedDateForExcel;
                    var testDate = UsedDateForExcel;

                    double DaysInCurrentPeriod = 0;
                    if (TabularType == "Day")
                    {
                        DaysInCurrentPeriod = 1;
                        begining = end = UsedDateForExcel;
                    }
                    else if (TabularType == "Week")
                    {
                        GetWeek(testDate, new CultureInfo("fr-FR"), out begining, out end); //en-US(Sunday - Monday) //fr-FR(Monday - Sunday)
                        if (end.Subtract(toDate).TotalSeconds > 0)
                        {
                            end = toDate;
                        }
                        if (begining.Subtract(UsedDateForExcel).TotalSeconds < 0)
                        {
                            begining = UsedDateForExcel;
                        }
                        DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                    }
                    else if (TabularType == "Month")
                    {
                        DateTime itempDate = UsedDateForExcel.AddMonths(1);
                        DateTime Temp2 = new DateTime(itempDate.Year, itempDate.Month, 01, 00, 00, 01);
                        end = Temp2.AddDays(-1);
                        if (end.Subtract(toDate).TotalSeconds > 0)
                        {
                            end = toDate;
                        }
                        DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                    }
                    else if (TabularType == "Year")
                    {
                        DateTime Temp2 = new DateTime(UsedDateForExcel.AddYears(1).Year, 01, 01, 00, 00, 01);
                        end = Temp2.AddDays(-1);
                        if (end.Subtract(toDate).TotalSeconds > 0)
                        {
                            end = toDate;
                        }
                        DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                    }

                    int StartingRowForToday = Row;
                    double IndividualDateOpTime = 0, IndividualDateSCTvsPP = 0, IndividualDateScrapQtyTime = 0, IndividualDateReWorkTime = 0;
                    double IndividualDateSetting = 0, IndividualDateIdle = 0, IndividualDateMinorLoss = 0, IndividualDateBreakdown = 0, IndividualDateNoPlan = 0;
                    string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");

                    double AvailableTimePerDay = 0;
                    int NumMacsToExcel = 0;


                    for (int n = 0; n < machin.Rows.Count; n++)
                    {
                        NumMacsToExcel++;
                        double CummulativeOfAllLosses = 0;
                        if (n == 0 && l != 0)
                        {
                            Row++;
                            StartingRowForToday = Row;
                        }

                        int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                        List<string> HierarchyData = GetHierarchyData(MachineID);

                        double AvaillabilityFactor = 0, EfficiencyFactor = 0, QualityFactor = 0;
                        double green, red, yellow, blue = 0, MinorLoss = 0, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
                        double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
                        double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;
                        double selfInspection = 0, Idle = 0;

                        //New Logic . Take values from 
                        string correctedDateS = UsedDateForExcel.ToString("yyyy-MM-dd");
                        string correctedDateE = UsedDateForExcel.ToString("yyyy-MM-dd");
                        if (TabularType != "Day")
                        {
                            if (end.Subtract(toDate).TotalSeconds < 0)
                            {
                                correctedDateE = end.ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                correctedDateE = toDate.ToString("yyyy-MM-dd");
                            }
                        }

                        //if (TabularType == "Day")
                        //{
                        //    //Default same as correctedDateE.
                        //}
                        //else if (TabularType == "Week")
                        //{
                        //    correctedDateE = end.ToString("yyyy-MM-dd");
                        //}
                        //else if (TabularType == "Month")
                        //{
                        //    correctedDateE = end.ToString("yyyy-MM-dd");
                        //}
                        //else if (TabularType == "Year")
                        //{
                        //    correctedDateE = end.ToString("yyyy-MM-dd");
                        //}

                        DateTime DateTimeValue = Convert.ToDateTime(UsedDateForExcel.ToString("yyyy-MM-dd") + " " + "00:00:00");
                        using (i_facility_talEntities1 dboee = new i_facility_talEntities1())
                        {
                            #region OLD 2017-03-30
                            //if (ProdFAI == "OverAll")
                            //{
                            //    var OEEData = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //    if (TabularType == "Day")
                            //    {
                            //        OEEData = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //    }
                            //    else if (TabularType == "Week")
                            //    {
                            //        OEEData = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //    }
                            //    else if (TabularType == "Month")
                            //    {
                            //        OEEData = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //    }
                            //    else if (TabularType == "Year")
                            //    {
                            //        OEEData = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //    }
                            //    if (OEEData != null)
                            //    {
                            //        MinorLosses = Convert.ToDouble(OEEData.MinorLosses);
                            //        blue = Convert.ToDouble(OEEData.Blue);
                            //        SettingTime = Convert.ToDouble(OEEData.SettingTime);
                            //        ROALossess = Convert.ToDouble(OEEData.ROALossess);
                            //        DownTimeBreakdown = Convert.ToDouble(OEEData.DownTimeBreakdown);
                            //        SummationOfSCTvsPP = Convert.ToDouble(OEEData.SummationOfSCTvsPP);
                            //        green = Convert.ToDouble(OEEData.Green);
                            //        OperatingTime = green;
                            //        ScrapQtyTime = Convert.ToDouble(OEEData.ScrapQtyTime);
                            //        ReWOTime = Convert.ToDouble(OEEData.ReWOTime);
                            //    }
                            //    else
                            //    {
                            //        #region Trying to Run .exe file with Params
                            //        try
                            //        {
                            //            String cPath = dboee.tblapp_paths.Where(m => m.IsDeleted == 0 && m.AppName == "CalOEEDaily").Select(m => m.AppPath).FirstOrDefault();
                            //            string filename = Path.Combine(cPath, "CalOEEDaily.exe");
                            //            var proc = System.Diagnostics.Process.Start(Server.MapPath(@filename), DateTimeValue.ToString());
                            //            proc.WaitForExit();
                            //            //proc.Kill();
                            //        }
                            //        catch (Exception e)
                            //        {
                            //        }
                            //        var OEEDataInner = dboee.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                            //        if (OEEDataInner != null)
                            //        {
                            //            MinorLosses = Convert.ToDouble(OEEDataInner.MinorLosses);
                            //            blue = Convert.ToDouble(OEEDataInner.Blue);
                            //            SettingTime = Convert.ToDouble(OEEDataInner.SettingTime);
                            //            ROALossess = Convert.ToDouble(OEEDataInner.ROALossess);
                            //            DownTimeBreakdown = Convert.ToDouble(OEEDataInner.DownTimeBreakdown);
                            //            SummationOfSCTvsPP = Convert.ToDouble(OEEData.SummationOfSCTvsPP);
                            //            green = Convert.ToDouble(OEEDataInner.Green);
                            //            OperatingTime = green;
                            //            ScrapQtyTime = Convert.ToDouble(OEEDataInner.ScrapQtyTime);
                            //            ReWOTime = Convert.ToDouble(OEEDataInner.ReWOTime);
                            //        }
                            //        else
                            //        {
                            //            continue;
                            //        }
                            //        #endregion
                            //    }
                            //}
                            #endregion

                            if (ProdFAI == "OverAll")
                            {
                                DataTable OEEData = new DataTable();
                                using (MsqlConnection mcOEE = new MsqlConnection())
                                {
                                    mcOEE.open();
                                    string query = null;
                                    query = @"select WCID,sum(Green), sum(SummationOfSCTvsPP),sum(SettingTime),sum(ROALossess),SUM(MinorLosses), sum(Blue), sum(DownTimeBreakdown), "
                                            + " sum(ScrapQtyTime),sum(ReWOTime) from i_facility_tal.dbo.tbloeedashboardvariables where StartDate >= '" + correctedDateS + "' and StartDate <= '" + correctedDateE + "' and WCID = '" + MachineID + "' group by WCID  ;";

                                    //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                                    SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                                    daLossCodesData.Fill(OEEData);
                                    mcOEE.close();
                                    //IntoFile(query);
                                }
                                if (OEEData.Rows.Count > 0)
                                {
                                    green = Convert.ToDouble(OEEData.Rows[0][1]);
                                    SummationOfSCTvsPP = Convert.ToDouble(OEEData.Rows[0][2]);
                                    SettingTime = Convert.ToDouble(OEEData.Rows[0][3]);
                                    //selfInspection = Convert.ToDouble(OEEData.Rows[0][4]); ;
                                    Idle = Convert.ToDouble(OEEData.Rows[0][4]);
                                    MinorLosses = Convert.ToDouble(OEEData.Rows[0][5]);
                                    ROALossess = selfInspection + Idle;
                                    blue = Convert.ToDouble(OEEData.Rows[0][6]);
                                    DownTimeBreakdown = Convert.ToDouble(OEEData.Rows[0][7]);
                                    ScrapQtyTime = Convert.ToDouble(OEEData.Rows[0][8]);
                                    ReWOTime = Convert.ToDouble(OEEData.Rows[0][9]);
                                    OperatingTime = green;
                                }
                            }
                            else
                            {
                                //var OEEData = dboee.tblworeports.Where(m => m.CorrectedDate == correctedDate && m.MachineID == MachineID && m.Type == ProdFAI).GroupBy(m=>m.MachineID).Select new { MachineID = };
                                DataTable OEEData = new DataTable();
                                using (MsqlConnection mcOEE = new MsqlConnection())
                                {
                                    mcOEE.open();
                                    string query = null;
                                    if (TabularType == "Day")
                                    {
                                        if (ProdFAI != "Others")
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "' group by MachineID;";
                                        }
                                        else
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "'  and  Type != 'FAI' and Type != 'Prod' group by MachineID;";
                                        }
                                    }
                                    else if (TabularType == "Week")
                                    {
                                        if (ProdFAI != "Others")
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "' group by MachineID;";
                                        }
                                        else
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "'  and  Type != 'FAI' and Type != 'Prod' group by MachineID;";
                                        }
                                    }
                                    else if (TabularType == "Month")
                                    {
                                        if (ProdFAI != "Others")
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "'  group by MachineID;";
                                        }
                                        else
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "'  and  Type != 'FAI' and Type != 'Prod'   group by MachineID;";
                                        }
                                    }
                                    else if (TabularType == "Year")
                                    {
                                        if (ProdFAI != "Others")
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "'  group by MachineID;";
                                        }
                                        else
                                        {
                                            query = @"select MachineID,sum(CuttingTime), sum(SummationOfSCTvsPP),sum(SettingTime),sum(SelfInspection),sum(Idle),SUM(MinorLoss), sum(Blue), sum(Breakdown),sum(ScrapQtyTime),sum(ReWorkTime) "
                                                         + " from i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "'  and  Type != 'FAI' and Type != 'Prod'   group by MachineID;";
                                        }
                                    }

                                    //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                                    SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                                    daLossCodesData.Fill(OEEData);
                                    mcOEE.close();
                                }
                                if (OEEData.Rows.Count > 0)
                                {
                                    green = Convert.ToDouble(OEEData.Rows[0][1]);
                                    SummationOfSCTvsPP = Convert.ToDouble(OEEData.Rows[0][2]);
                                    SettingTime = Convert.ToDouble(OEEData.Rows[0][3]);
                                    selfInspection = Convert.ToDouble(OEEData.Rows[0][4]); ;
                                    Idle = Convert.ToDouble(OEEData.Rows[0][5]); ;
                                    MinorLosses = Convert.ToDouble(OEEData.Rows[0][6]);
                                    ROALossess = selfInspection + Idle;
                                    blue = Convert.ToDouble(OEEData.Rows[0][7]);
                                    DownTimeBreakdown = Convert.ToDouble(OEEData.Rows[0][8]);
                                    ScrapQtyTime = Convert.ToDouble(OEEData.Rows[0][9]);
                                    ReWOTime = Convert.ToDouble(OEEData.Rows[0][10]);
                                    OperatingTime = green;
                                }
                            }
                        }

                        // OperatingTime = AvailableTime - (ROALossess + DownTimeBreakdown + blue + MinorLosses + ROPLosses + ROQLosses);
                        // OperatingTime = AvailableTime - (ROALossess + DownTimeBreakdown + blue + MinorLosses);

                        setup = SettingTime;
                        //double TotalMinutes = green + setup + (yellow - setup) + red + blue;
                        double TotalMinutes = OperatingTime + SettingTime + (ROALossess - SettingTime) + MinorLosses + DownTimeBreakdown + blue;
                        //double Diff = 1440 - TotalMinutes;

                        //for 3 shifts so 
                        //double Diff = (8 * 3 * 60) - TotalMinutes;
                        Double ActualTotalMinutes = (8 * 3 * 60) * DaysInCurrentPeriod;
                        double Diff = ActualTotalMinutes - TotalMinutes;
                        if (Diff > 0)
                        {
                            blue += Diff;
                        }
                        else
                        {
                            ROALossess += Diff;
                        }

                        if (TimeFactor == "GH")
                        {
                            AvailableTime = ActualTotalMinutes; //24Hours to Minutes
                        }
                        else if (TimeFactor == "NoBlue")
                        {
                            AvailableTime = ActualTotalMinutes - blue;
                        }
                        worksheet.Cells["Q" + Row].Value = Math.Round(AvailableTime / 60, 2); 

                        int IsWorking = 1;
                        AvailableTimePerDay += AvailableTime;

                        worksheet.Cells["B" + Row].Value = Sno++;
                        //worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                        worksheet.Cells["C" + Row].Value = HierarchyData[0];
                        worksheet.Cells["D" + Row].Value = HierarchyData[1];
                        worksheet.Cells["E" + Row].Value = HierarchyData[2];
                        worksheet.Cells["F" + Row].Value = HierarchyData[4];
                        worksheet.Cells["G" + Row].Value = HierarchyData[3];

                        worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                        worksheet.Cells["I" + Row].Value = end.ToString("yyyy-MM-dd");

                        #region Calculate and push A,P,Q and OEE

                        //Utilisation // Whole day duration is 24*60 = 1440 minutes
                        double valA = 0;
                        valA = OperatingTime / AvailableTime;

                        //Availablity
                        if (valA > 0 && valA < 100)
                        {
                            worksheet.Cells["J" + Row].Value = Math.Round(valA * 100, 0);
                            AvaillabilityFactor = Math.Round(valA * 100, 0);
                        }
                        else
                        {
                            worksheet.Cells["J" + Row].Value = 0;
                            AvaillabilityFactor = 0;
                        }

                        if (AvaillabilityFactor != 0)
                        {
                            //Performance
                            if (SummationOfSCTvsPP == -1 || SummationOfSCTvsPP == 0)
                            {
                                EfficiencyFactor = 100;
                                worksheet.Cells["K" + Row].Value = Math.Round(EfficiencyFactor, 0);
                            }
                            else
                            {
                                EfficiencyFactor = Math.Round((SummationOfSCTvsPP / (OperatingTime)) * 100, 0);
                                if (EfficiencyFactor >= 0 && EfficiencyFactor <= 100)
                                    worksheet.Cells["K" + Row].Value = EfficiencyFactor;
                                else if (EfficiencyFactor > 100)
                                {
                                    EfficiencyFactor = 100;
                                    worksheet.Cells["K" + Row].Value = 100;
                                }
                                else if (EfficiencyFactor < 0)
                                {
                                    EfficiencyFactor = 0;
                                    worksheet.Cells["K" + Row].Value = 0;
                                }
                            }
                            //Quality
                            if (OperatingTime != 0)
                            {
                                QualityFactor = ((OperatingTime - ScrapQtyTime - ReWOTime) / OperatingTime) * 100;
                                if (QualityFactor >= 0 && QualityFactor <= 100)
                                {
                                    worksheet.Cells["L" + Row].Value = Math.Round(QualityFactor, 0);
                                }
                                else if (QualityFactor > 100)
                                {
                                    QualityFactor = 100;
                                    worksheet.Cells["L" + Row].Value = 100;
                                }
                                else if (QualityFactor < 0)
                                {
                                    QualityFactor = 0;
                                    worksheet.Cells["L" + Row].Value = 0;
                                }
                            }
                            else
                            {
                                QualityFactor = 0;
                                worksheet.Cells["L" + Row].Value = 0;
                            }
                        }
                        else
                        {
                            worksheet.Cells["K" + Row].Value = 0;//Performance
                            worksheet.Cells["L" + Row].Value = 0;//Quality
                        }

                        //OEE
                        if (AvaillabilityFactor <= 0 || EfficiencyFactor <= 0 || QualityFactor <= 0)
                        {
                            worksheet.Cells["M" + Row].Value = 0;
                        }
                        else
                        {
                            valA = Math.Round((AvaillabilityFactor / 100) * (EfficiencyFactor / 100) * (QualityFactor / 100) * 100, 0);
                            if (valA >= 0 && valA <= 100)
                            {
                                worksheet.Cells["M" + Row].Value = Math.Round(valA, 0);
                            }
                            else if (valA > 100)
                            {
                                worksheet.Cells["M" + Row].Value = 100;
                            }
                            else if (valA < 0)
                            {
                                worksheet.Cells["M" + Row].Value = 0;
                            }
                        }

                        #endregion

                        //if (TimeFactor == "GH")
                        //{
                        //    worksheet.Cells["N" + Row].Value = 24;
                        //}

                        worksheet.Cells["N" + Row].Value = Math.Round((SummationOfSCTvsPP / 60), 1);

                        //worksheet.Cells["R" + Row].Value = 24 * 60;
                        worksheet.Cells["S" + Row].Formula = "=SUM(R" + Row + ",T" + Row + ",U" + Row + ",V" + Row + ")";
                        worksheet.Cells["R" + Row].Value = Math.Round(OperatingTime / 60, 1);

                        worksheet.Cells["P" + Row].Value = 1;

                        //To push Formula for Total,a column @ index(20) in Template
                        //string LossesEndsAtCol = ExcelColumnFromNumber(20 + LossesList.Count);
                        //worksheet.Cells["T" + Row].Formula = "=SUM(U" + Row + ":" + LossesEndsAtCol + "" + Row + ")";

                        double ValueAddingTime = Math.Round(OperatingTime, 2);
                        double setTime = Math.Round(SettingTime, 2);
                        double idleTime = Math.Round((ROALossess - SettingTime), 2);
                        double minorLossTime = Math.Round(MinorLosses, 2);
                        double BreakdownTime = Math.Round(DownTimeBreakdown, 2);
                        double blueTime = Math.Round(blue, 2);

                        //For Individual Date Cummulative
                        IndividualDateOpTime += ValueAddingTime;
                        IndividualDateSCTvsPP += SummationOfSCTvsPP;
                        IndividualDateScrapQtyTime += ScrapQtyTime;
                        IndividualDateReWorkTime += ReWOTime;

                        IndividualDateSetting += setTime;
                        IndividualDateIdle += idleTime;
                        IndividualDateMinorLoss += minorLossTime;
                        IndividualDateBreakdown += BreakdownTime;
                        IndividualDateNoPlan += blueTime;

                        //Added this machineDetails into Datatable
                        string WCInvNoString = HierarchyData[3];
                        //DataRow dr = DTLosses.Select("LossName= " + lossname).FirstOrDefault(); 
                        //DataRow dr = DTConsolidatedLosses.Select("WCInvNo = '" + @WCInvNoString + "'", " CorrectedDate= '" + @dateforMachine + "'").FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                        DataRow dr = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoString && r.Field<string>("CorrectedDate") == dateforMachine);
                        if (dr != null)
                        {
                            //do nothing
                        }
                        else
                        {
                            //plant, shop, cell, macINV, WcName, CorrectedDate, ValueAdding(Green/Operating), AvailableTime, SummationofSCTvsPP, Scrap,Rework,CuttingTime,DaysWorking, GodHours, TotalSTDHours, RejectionHours.
                            DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3], HierarchyData[4], dateforMachine, ValueAddingTime, AvailableTime, SummationOfSCTvsPP, ScrapQtyTime, ReWOTime, ValueAddingTime, IsWorking, AvailableTime, 24, 0, minorLossTime, BreakdownTime);
                        }
                        //Now get & put Losses
                        // Push Loss Value into  DataTable & Excel

                        //1st push 0 for every loss into excel
                        int column = 22 + LossCodesData.Rows.Count; // StartCol in Excel + TotalLosses
                        finalLossCol = ExcelColumnFromNumber(column);
                        worksheet.Cells["W" + Row + ":" + finalLossCol + "" + Row].Value = 0; ;

                        #region
                        if (ProdFAI == "OverAll")
                        {
                            //to Capture and Push , Losses that occured.
                            List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSecondsForDateRange(MachineID, correctedDateS, correctedDateE);
                            DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                            if (dr1 != null)
                            {
                                foreach (var loss in LossesdurationList)
                                {
                                    int LossID = loss.Key;
                                    double Duration = loss.Value;
                                    var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossID).FirstOrDefault();
                                    int level = lossdata.LossCodesLevel;
                                    string losscodeName = null;

                                    #region To Get LossCode Hierarchy
                                    if (level == 3)
                                    {
                                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                        int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                                        var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                                        var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                                        losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                    }
                                    else if (level == 2)
                                    {
                                        int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                        var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                                        losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                    }
                                    else if (level == 1)
                                    {
                                        if (LossID == 999)
                                        {
                                            losscodeName = "NoCode Entered";
                                        }
                                        else
                                        {
                                            losscodeName = lossdata.LossCode;
                                        }
                                    }
                                    #endregion

                                    int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                    string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 1);
                                    double DurInHours = Convert.ToDouble(Math.Round((Duration / (60 * 60)), 1)); //To Hours:: 1 Decimal Place
                                    worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                                    dr1[losscodeName] = DurInHours;
                                    CummulativeOfAllLosses += DurInHours;
                                }
                            }
                        }
                        else
                        {
                            DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                            if (dr1 != null)
                            {
                                DataTable LossesDurationData = new DataTable();
                                using (MsqlConnection mcLossCodes = new MsqlConnection())
                                {
                                    mcLossCodes.open();
                                    string query = null;
                                    if (ProdFAI != "Others")
                                    {
                                        query = @"select sum(LossDuration),Level,LossName,LossCodeLevel1Name,LossCodeLevel2Name from i_facility_tal.dbo.tblwolossess Where HMIID in 
                                            ( SELECT HMIID FROM i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "' ) group by LossName;";
                                    }
                                    else
                                    {
                                        query = @"select sum(LossDuration),Level,LossName,LossCodeLevel1Name,LossCodeLevel2Name from i_facility_tal.dbo.tblwolossess Where HMIID in 
                                            ( SELECT HMIID FROM i_facility_tal.dbo.tblworeport where CorrectedDate >= '" + correctedDateS + "' and CorrectedDate <= '" + correctedDateE + "' and MachineID = '" + MachineID + "' and  Type != 'FAI' and Type != 'Prod'  ) group by LossName;";

                                    }
                                    //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                                    SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                                    daLossCodesData.Fill(LossesDurationData);
                                    mcLossCodes.close();
                                }

                                for (int Lossloop = 0; Lossloop < LossesDurationData.Rows.Count; Lossloop++)
                                {
                                    double Duration = Convert.ToDouble(LossesDurationData.Rows[Lossloop][0]);
                                    int level = Convert.ToInt32(LossesDurationData.Rows[Lossloop][1]);
                                    string losscodeName = null;

                                    #region To Get LossCode Hierarchy
                                    if (level == 3)
                                    {
                                        string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][3]);
                                        string Level2Name = Convert.ToString(LossesDurationData.Rows[Lossloop][4]);
                                        string Level3Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                        losscodeName = Level1Name + " :: " + Level2Name + " : " + Level3Name;
                                    }
                                    else if (level == 2)
                                    {
                                        string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][3]);
                                        string Level2Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                        losscodeName = Level1Name + ":" + Level2Name;
                                    }
                                    else if (level == 1)
                                    {
                                        string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                        if (Level1Name == "999")
                                        {
                                            losscodeName = "NoCode Entered";
                                        }
                                        else
                                        {
                                            losscodeName = Level1Name;
                                        }
                                    }
                                    #endregion

                                    int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                    string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 1);
                                    double DurInHours = Convert.ToDouble(Math.Round((Duration / (60 * 60)), 1)); //To Hours:: 1 Decimal Place
                                    worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                                    dr1[losscodeName] = DurInHours;
                                    CummulativeOfAllLosses += DurInHours;
                                }
                            }
                        }
                        #endregion
                        worksheet.Cells["T" + Row].Value = Convert.ToDouble(Math.Round((CummulativeOfAllLosses), 1));
                        worksheet.Cells["U" + Row].Value = Math.Round(minorLossTime / 60, 2);
                        worksheet.Cells["V" + Row].Value = Math.Round(BreakdownTime / 60, 2);
                        Row++;
                    }

                    #region //Daywise OEE and Stuff
                    //1) Availability
                    double AFactor = 0;
                    AFactor = IndividualDateOpTime / AvailableTimePerDay;
                    if (AFactor > 0)
                    {
                        worksheet.Cells["J" + Row].Value = Math.Round(AFactor * 100, 0);
                        AFactor = Math.Round(AFactor * 100, 0);
                    }
                    else
                    {
                        worksheet.Cells["J" + Row].Value = 0;
                        AFactor = 0;
                    }
                    double QFactor = 0; double EFactor = 0;
                    if (AFactor != 0)
                    {

                        //2)Performance
                        if (IndividualDateSCTvsPP <= 0 || IndividualDateOpTime <= 0)
                        {
                            EFactor = 100;
                            worksheet.Cells["K" + Row].Value = Math.Round(EFactor, 0);
                        }
                        else
                        {
                            EFactor = Math.Round((IndividualDateSCTvsPP / (IndividualDateOpTime)) * 100, 0);
                            if (EFactor > 0 && EFactor <= 100)
                            {
                                worksheet.Cells["K" + Row].Value = EFactor;
                            }
                            else if (EFactor > 100)
                            {
                                EFactor = 100;
                                worksheet.Cells["K" + Row].Value = 100;
                            }
                            else if (EFactor <= 0)
                            {
                                EFactor = 0;
                                worksheet.Cells["K" + Row].Value = 100;
                            }
                        }
                        //3) Quality
                        if (IndividualDateOpTime != 0)
                        {
                            QFactor = ((IndividualDateOpTime - IndividualDateScrapQtyTime - IndividualDateReWorkTime) / IndividualDateOpTime) * 100;
                            if (QFactor > 0 && QFactor <= 100)
                            {
                                worksheet.Cells["L" + Row].Value = Math.Round(QFactor, 0);
                                QFactor = Math.Round(QFactor, 2);
                            }
                            else if (QFactor > 100)
                            {
                                QFactor = 100;
                                worksheet.Cells["L" + Row].Value = 100;
                            }
                            else if (QFactor <= 0)
                            {
                                QFactor = 0;
                                worksheet.Cells["L" + Row].Value = 100;
                            }
                        }
                        else
                        {
                            QFactor = 100;
                            worksheet.Cells["L" + Row].Value = 100;
                        }

                    }
                    else
                    {
                        worksheet.Cells["K" + Row].Value = 0;//Performance
                        worksheet.Cells["L" + Row].Value = 0;//Quality
                    }


                    //4) OEE
                    double ValOEE = 0;
                    double OEEFactor = 0;
                    if (AFactor <= 0 || EFactor <= 0 || QFactor <= 0)
                    {
                        worksheet.Cells["M" + Row].Value = 0;
                    }
                    else
                    {
                        ValOEE = Math.Round(((AFactor / 100) * (EFactor / 100) * (QFactor / 100)) * 100, 0);
                        if (ValOEE >= 0 && ValOEE <= 100)
                        {
                            OEEFactor = Math.Round(ValOEE, 2);
                            worksheet.Cells["M" + Row].Value = Math.Round(ValOEE, 0);
                        }
                        else if (ValOEE > 100)
                        {
                            OEEFactor = 100;
                            worksheet.Cells["M" + Row].Value = 100;
                        }
                        else if (ValOEE < 0)
                        {
                            worksheet.Cells["M" + Row].Value = 0;
                        }
                    }

                    //OEE and  Stuff for entire day (of all WC's) Into DT
                    DTConsolidatedLosses.Rows.Add("Summarized", "Summarized", "Summarized", "Summarized", "Summarized", dateforMachine, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, "J" + Row, "K" + Row, "L" + Row, "M" + Row);

                    //Cellwise Border for Today
                    worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Insert Cummulative for today
                    worksheet.Cells["C" + Row + ":G" + Row].Merge = true;
                    worksheet.Cells["C" + Row].Value = "Summarized For";
                    worksheet.Cells["H" + Row].Value = worksheet.Cells["H" + (Row - 1)].Value;
                    worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Font.Bold = true;
                    //worksheet.Cells["I" + Row].Value = worksheet.Cells["I" + (Row - 1)].Value;

                    // 1) Total of AllLosses( Col : T )
                    worksheet.Cells["T" + Row].Formula = "=SUM(T" + StartingRowForToday + ":T" + (Row - 1) + ")";
                    // 2) Total of OperatingTime( Col : S )
                    worksheet.Cells["S" + Row].Formula = "=SUM(S" + StartingRowForToday + ":S" + (Row - 1) + ")";
                    // 2) Total of CuttingTime( Col : R )
                    worksheet.Cells["R" + Row].Formula = "=SUM(R" + StartingRowForToday + ":R" + (Row - 1) + ")";
                    // 3) God Hours( Col : Q )
                    worksheet.Cells["Q" + Row].Formula = "=SUM(Q" + StartingRowForToday + ":Q" + (Row - 1) + ")";
                    // 4) Days Working( Col : P )
                    worksheet.Cells["P" + Row].Formula = "=SUM(P" + StartingRowForToday + ":P" + (Row - 1) + ")";
                    // 5) Days Working( Col : N )
                    worksheet.Cells["N" + Row].Formula = "=SUM(N" + StartingRowForToday + ":N" + (Row - 1) + ")";
                    // 6) MinorLoss( Col : U )
                    worksheet.Cells["U" + Row].Formula = "=SUM(U" + StartingRowForToday + ":U" + (Row - 1) + ")";
                    // 7) Breakdown( Col : V )
                    worksheet.Cells["V" + Row].Formula = "=SUM(V" + StartingRowForToday + ":V" + (Row - 1) + ")";

                    //5) Border for Above 4 & Around them.
                    //worksheet.Cells["P" + StartingRowForToday + ":S" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    //Push each Date Cummulative. Loop through ExcelAddress and insert formula
                    var rangeIndividualSummarized = worksheet.Cells["W" + Row + ":" + finalLossCol + "" + Row];
                    foreach (var rangeBase in rangeIndividualSummarized)
                    {
                        string str = Convert.ToString(rangeBase);
                        string ExcelColAlphabet = Regex.Replace(str, "[^A-Z _]", "");
                        worksheet.Cells[rangeBase.Address].Formula = "=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")";
                        //var a = worksheet.Cells[rangeBase.Address].Value;
                        var blah1 = worksheet.Calculate("=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")");

                        double LossVal = 0;
                        double.TryParse(Convert.ToString(blah1), out LossVal);
                        if (LossVal != 0.0)
                        {
                            string LossName = Convert.ToString(worksheet.Cells[ExcelColAlphabet + 3].Value);
                            DataRow dr = DTConsolidatedLosses.AsEnumerable().LastOrDefault(r => r.Field<string>("Plant") == "Summarized" && r.Field<string>("CorrectedDate") == dateforMachine);
                            if (dr != null)
                            {
                                dr[LossName] = LossVal;
                            }
                        }
                    }

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    #endregion

                    //UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                    double offsetE = 1;
                    if (TabularType == "Day")
                    {
                        offsetE = 1;
                    }
                    else if (TabularType == "Week")
                    {
                        offsetE = end.Subtract(begining).TotalDays + 1;
                    }
                    else if (TabularType == "Month")
                    {
                        offsetE = end.Subtract(begining).TotalDays + 1;
                    }
                    else if (TabularType == "Year")
                    {
                        offsetE = end.Subtract(begining).TotalDays + 1;
                    }

                    UsedDateForExcel = UsedDateForExcel.AddDays(offsetE);
                    Row++;
                    l++;
                } while (UsedDateForExcel <= toDate && toDate.Date != DateTime.Now.Date);

                #region OverAll OEE and Stuff

                Row = 5;
                Sno = 1;
                var WCInvNoList = (from DataRow row in DTConsolidatedLosses.Rows
                                   where row["WCInvNo"] != "Summarized"
                                   select row["WCInvNo"]).Distinct();

                double OverAllOpTime = 0, OverAllAvailableTime = 0, OverAllSCTvsPP = 0, OverAllScrapQtyTime = 0, OverAllReworkTime = 0, OverAllMinorLoss = 0, OverAllBreakdown = 0;
                foreach (var MacINV in WCInvNoList)
                {
                    string WCInvNoStringOverAll = Convert.ToString(MacINV);
                    DataRow drOverAll = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoStringOverAll);

                    if (drOverAll != null)
                    {
                        int MachineID = db.tblmachinedetails.Where(m => m.MachineInvNo == WCInvNoStringOverAll).Select(m => m.MachineID).SingleOrDefault();
                        string macDispName = db.tblmachinedetails.Where(m => m.MachineInvNo == WCInvNoStringOverAll).Select(m => m.MachineDispName).SingleOrDefault();
                        List<string> HierarchyData = GetHierarchyData(MachineID);
                        worksheet.Cells["B" + Row].Value = Sno++;
                        worksheet.Cells["C" + Row].Value = HierarchyData[0];
                        worksheet.Cells["D" + Row].Value = HierarchyData[1];
                        worksheet.Cells["E" + Row].Value = HierarchyData[2];
                        worksheet.Cells["F" + Row].Value = macDispName;
                        worksheet.Cells["G" + Row].Value = HierarchyData[3];

                        worksheet.Cells["H" + Row].Value = (frmDate).ToString("yyyy-MM-dd");
                        worksheet.Cells["I" + Row].Value = (toDate).ToString("yyyy-MM-dd");

                        //OEE and Stuff
                        double OpTime = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("OpTime"));
                        double AvailableTime = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("AvailableTime"));
                        double SCTvsPP = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("SCTvsPP"));
                        double ScrapQtyTime = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("ScrapQtyTime"));
                        double ReworkTime = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("ReworkTime"));
                        double CuttingTime = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("CuttingTime"));
                        int DaysWorking = (int)DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("DaysWorking"));
                        int GodHours = (int)DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("GodHours"));
                        double TotalSTDHours = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("TotalSTDHours"));
                        double RejectionHours = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("RejectionHours"));
                        double MinorLoss = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("MinorLoss"));
                        double Breakdown = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>("Breakdown"));

                        OverAllOpTime += OpTime;
                        OverAllAvailableTime += AvailableTime;
                        OverAllReworkTime += ReworkTime;
                        OverAllScrapQtyTime += ScrapQtyTime;
                        OverAllSCTvsPP += SCTvsPP;
                        OverAllMinorLoss += MinorLoss;
                        OverAllBreakdown += Breakdown;

                        //Value InTo 
                        //worksheet.Cells["S" + Row].Value = Math.Round( OpTime, 2);
                        worksheet.Cells["R" + Row].Value = Math.Round((OpTime / 60), 1);
                        worksheet.Cells["S" + Row].Formula = "=SUM(R" + Row + ",T" + Row + ")";
                        worksheet.Cells["Q" + Row].Value = Math.Round(Convert.ToDouble(GodHours) / 60, 2);
                        worksheet.Cells["P" + Row].Value = DaysWorking;
                        worksheet.Cells["O" + Row].Value = Math.Round(RejectionHours, 1);
                        worksheet.Cells["N" + Row].Value = Math.Round((SCTvsPP / 60), 1);
                        worksheet.Cells["U" + Row].Value = Math.Round(MinorLoss / 60, 1);
                        worksheet.Cells["V" + Row].Value = Math.Round((Breakdown / 60), 1);

                        #region A,E,Q & OEE
                        double AvaillabilityFactor = 0, EfficiencyFactor = 0, QualityFactor = 0;

                        double valA = 0;
                        valA = OpTime / AvailableTime;

                        //Availablity
                        if (valA > 0)
                        {
                            worksheet.Cells["J" + Row].Value = Math.Round(valA * 100, 0);
                            AvaillabilityFactor = Math.Round(valA * 100, 0);
                        }
                        else
                        {
                            worksheet.Cells["J" + Row].Value = 0;
                            AvaillabilityFactor = 0;
                        }
                        if (AvaillabilityFactor != 0)
                        {
                            //Performance
                            if (SCTvsPP == -1 || SCTvsPP == 0)
                            {
                                EfficiencyFactor = 100;
                                worksheet.Cells["K" + Row].Value = Math.Round(EfficiencyFactor, 0);
                            }
                            else
                            {
                                EfficiencyFactor = Math.Round((SCTvsPP / (OpTime)) * 100, 0);
                                if (EfficiencyFactor > 0 && EfficiencyFactor <= 100)
                                    worksheet.Cells["K" + Row].Value = EfficiencyFactor;
                                else if (EfficiencyFactor > 100)
                                {
                                    EfficiencyFactor = 100;
                                    worksheet.Cells["K" + Row].Value = 100;
                                }
                                else if (EfficiencyFactor <= 0)
                                {
                                    EfficiencyFactor = 0;
                                    worksheet.Cells["K" + Row].Value = 100;
                                }
                            }

                            //Quality
                            if (OpTime != 0)
                            {
                                QualityFactor = Math.Round(((OpTime - ScrapQtyTime - ReworkTime) / OpTime) * 100, 0);
                                if (QualityFactor > 0 && QualityFactor <= 100)
                                {
                                    worksheet.Cells["L" + Row].Value = QualityFactor;
                                }
                                else if (QualityFactor > 100)
                                {
                                    QualityFactor = 100;
                                    worksheet.Cells["L" + Row].Value = 100;
                                }
                                else if (QualityFactor <= 0)
                                {
                                    QualityFactor = 0;
                                    worksheet.Cells["L" + Row].Value = 100;
                                }
                            }
                            else
                            {
                                QualityFactor = 0;
                                worksheet.Cells["L" + Row].Value = 0;
                            }

                        }
                        //###################
                        else
                        {
                            worksheet.Cells["K" + Row].Value = 0;//Performance
                            worksheet.Cells["L" + Row].Value = 0;//Quality
                        }


                        //OEE
                        if (AvaillabilityFactor <= 0 || EfficiencyFactor <= 0 || QualityFactor <= 0)
                        {
                            worksheet.Cells["M" + Row].Value = 0;
                        }
                        else
                        {
                            valA = Math.Round((AvaillabilityFactor / 100) * (EfficiencyFactor / 100) * (QualityFactor / 100) * 100, 2);
                            if (valA >= 0 && valA <= 100)
                            {
                                worksheet.Cells["M" + Row].Value = Math.Round(valA, 0);
                            }
                            else if (valA > 100)
                            {
                                worksheet.Cells["M" + Row].Value = 100;
                            }
                            else if (valA < 0)
                            {
                                worksheet.Cells["M" + Row].Value = 0;
                            }
                        }
                        #endregion

                        //Total of Losses
                        worksheet.Cells["T" + Row].Formula = "=SUM(W" + Row + ":" + finalLossCol + "" + Row + ")";

                        //OverAll Losses 
                        var range = worksheet.Cells["W" + 3 + ":" + finalLossCol + "" + 3];
                        int i = 23;

                        foreach (var rangeBase in range)
                        {
                            string LossNameVal = Convert.ToString(rangeBase.Value);
                            string LossName = Convert.ToString(worksheet.Cells[rangeBase.Address].Value);
                            double LossValToExcel = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>(@LossNameVal));
                            string ColumnForThisLoss = ExcelColumnFromNumber(i++);
                            worksheet.Cells[ColumnForThisLoss + "" + Row].Value = Math.Round(LossValToExcel, 2);
                        }
                    }
                    Row++;
                }

                //Now Calculate OEE and Stuff with OverAll Variables calculated above "OverAll is the keywork in variable name"

                #region A,E,Q & OEE
                double OverAllAvaillabilityFactor = 0, OverAllEfficiencyFactor = 0, OverAllQualityFactor = 0;

                double OverAllvalA = 0;
                OverAllvalA = OverAllOpTime / OverAllAvailableTime;

                //Availablity
                if (OverAllvalA > 0)
                {
                    worksheet.Cells["J" + Row].Value = Math.Round(OverAllvalA * 100, 0);
                    OverAllAvaillabilityFactor = Math.Round(OverAllvalA * 100, 0);
                }
                else
                {
                    worksheet.Cells["J" + Row].Value = 0;
                    OverAllAvaillabilityFactor = 0;
                }
                if (OverAllAvaillabilityFactor != 0)
                {
                    //Performance
                    if (OverAllSCTvsPP == -1 || OverAllSCTvsPP == 0)
                    {
                        OverAllEfficiencyFactor = 100;
                        worksheet.Cells["K" + Row].Value = Math.Round(OverAllEfficiencyFactor, 0);
                    }
                    else
                    {
                        OverAllEfficiencyFactor = Math.Round((OverAllSCTvsPP / (OverAllOpTime)) * 100, 0);
                        if (OverAllEfficiencyFactor > 0 && OverAllEfficiencyFactor <= 100)
                            worksheet.Cells["K" + Row].Value = OverAllEfficiencyFactor;
                        else if (OverAllEfficiencyFactor > 100)
                        {
                            OverAllEfficiencyFactor = 100;
                            worksheet.Cells["K" + Row].Value = 100;
                        }
                        else if (OverAllEfficiencyFactor <= 0)
                        {
                            OverAllEfficiencyFactor = 0;
                            worksheet.Cells["K" + Row].Value = 100;
                        }
                    }

                    //Quality
                    if (OverAllOpTime != 0)
                    {
                        OverAllQualityFactor = Math.Round(((OverAllOpTime - OverAllScrapQtyTime - OverAllReworkTime) / OverAllOpTime) * 100, 0);
                        if (OverAllQualityFactor > 0 && OverAllQualityFactor <= 100)
                        {
                            worksheet.Cells["L" + Row].Value = OverAllQualityFactor;
                        }
                        else if (OverAllQualityFactor > 100)
                        {
                            OverAllQualityFactor = 100;
                            worksheet.Cells["L" + Row].Value = 100;
                        }
                        else if (OverAllQualityFactor <= 0)
                        {
                            OverAllQualityFactor = 0;
                            worksheet.Cells["L" + Row].Value = 100;
                        }
                    }
                    else
                    {
                        OverAllQualityFactor = 0;
                        worksheet.Cells["L" + Row].Value = 0;
                    }

                }
                //###################
                else
                {
                    worksheet.Cells["K" + Row].Value = 0;//Performance
                    worksheet.Cells["L" + Row].Value = 0;//Quality
                }


                //OEE
                if (OverAllAvaillabilityFactor <= 0 || OverAllEfficiencyFactor <= 0 || OverAllQualityFactor <= 0)
                {
                    worksheet.Cells["M" + Row].Value = 0;
                }
                else
                {
                    OverAllvalA = Math.Round((OverAllAvaillabilityFactor / 100) * (OverAllEfficiencyFactor / 100) * (OverAllQualityFactor / 100) * 100, 2);
                    if (OverAllvalA >= 0 && OverAllvalA <= 100)
                    {
                        worksheet.Cells["M" + Row].Value = Math.Round(OverAllvalA, 0);
                    }
                    else if (OverAllvalA > 100)
                    {
                        worksheet.Cells["M" + Row].Value = 100;
                    }
                    else if (OverAllvalA < 0)
                    {
                        worksheet.Cells["M" + Row].Value = 0;
                    }
                }
                #endregion

                //Formulas to Calculate OverAll Summarized for Column N-T.
                int firstRowOfSummarizedOverAll = 5;
                int lastRowOfSummarizedOverAll = (4 + machin.Rows.Count);
                var rangeSummarizedOverall = worksheet.Cells["N" + Row + ":T" + Row];
                foreach (var rangeBase in rangeSummarizedOverall)
                {
                    string column = Regex.Replace(rangeBase.Address, @"[\d-]", string.Empty);
                    worksheet.Cells[rangeBase.Address].Formula = "=SUM(" + column + 5 + ":" + column + "" + (Row - 1) + ")"; ;
                }

                // Borders and Stuff for Cummulative Data.
                //Cellwise Border for Today
                worksheet.Cells["B3:" + finalLossCol + "" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B3:" + finalLossCol + "" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B3:" + finalLossCol + "" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B3:" + finalLossCol + "" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                //Excel:: Border Around Cells.
                worksheet.Cells["B5:" + finalLossCol + "" + (Row)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                worksheet.Cells["B" + Row + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);


                ////Formula to Cummulate Losses
                //var rangeFinalLosses = worksheet.Cells["T" + Row + ":" + finalLossCol + "" + Row];
                //int j = 20;
                //foreach (var rangeBase in rangeFinalLosses)
                //{
                //    string ColumnForThisLoss = ExcelColumnFromNumber(j++);
                //    worksheet.Cells[ColumnForThisLoss + "" + Row].Formula = "=SUM(" + ColumnForThisLoss + 5 + ":" + ColumnForThisLoss + "" + (Row - 1) + ")";
                //}

                //Cummulative Losses into DT and Occured Losses into List and Identified and UnIdentified Losses
                //var rangeFinalLosses = worksheet.Cells["U" + Row + ":" + finalLossCol + "" + Row];
                var rangeFinalLosses = worksheet.Cells["W3:" + finalLossCol + "3"];
                List<KeyValuePair<string, double>> AllOccuredLosses = new List<KeyValuePair<string, double>>();
                int j = 21;
                double IdentifiedLoss = 0;
                double UnIdentifiedLoss = 0;
                foreach (var rangeBase in rangeFinalLosses)
                {
                    string LossName = Convert.ToString(rangeBase.Value);
                    string LossNameAddress = Convert.ToString(rangeBase.Address);

                    double thisLossValue = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("Plant") != "Summarized").Sum(x => x.Field<double>(@LossName));
                    //string thisLossValueString = TimeFromSeconds(thisLossValue) == "0D, 00:00:00" ? "0" : TimeFromSeconds(thisLossValue);
                    //TimeSpan time = TimeSpan.FromSeconds(thisLossValue);
                    //string str = time.ToString(@"hh\:mm\:ss");
                    string ColumnForThisLoss = ExcelColumnFromNumber(j++);
                    worksheet.Cells[ColumnForThisLoss + "" + Row].Formula = "=SUM(" + ColumnForThisLoss + 5 + ":" + ColumnForThisLoss + "" + (Row - 1) + ")";
                    // Double ValVal = (double) worksheet.Calculate("=SUM(" + ColumnForThisLoss + 5 + ":" + ColumnForThisLoss + "" + (Row - 1) + ")");
                    // worksheet.Cells[ColumnForThisLoss + "" + Row].Value = Math.Round(thisLossValue,1);
                    //worksheet.Cells[ColumnForThisLoss + "" + (Row + 1)].Value = Convert.ToString(thisLossValueString);
                    if (thisLossValue > 0)
                    {
                        if (LossName == "NoCode Entered" || LossName == "Unidentified Breakdown")
                        {
                            UnIdentifiedLoss += thisLossValue;
                        }
                        else
                        {
                            IdentifiedLoss += thisLossValue;
                        }
                        AllOccuredLosses.Add(new KeyValuePair<string, double>(LossNameAddress, Math.Round(thisLossValue, 1)));
                    }
                }

                #endregion

                #region GRAPHS
                //Create the chart

                if (machin.Rows.Count > 0)
                {
                    #region OEE and Stuff
                    int TotalSummarizedRows = 0;
                    TotalSummarizedRows = (machin.Rows.Count - 1); //-1 as data starts @ 5.

                    //Its Not MacINV its MacDescription
                    ExcelRange erLossesRangeMacInv = worksheet.Cells["F5:F" + (5 + TotalSummarizedRows)];
                    //OEE
                    ExcelChart chartOEE1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartOEE", eChartType.ColumnClustered);
                    var chartOEE = (ExcelBarChart)chartOEE1;
                    chartOEE.SetSize(300, 390);
                    chartOEE.SetPosition(60, 10);
                    chartOEE.Title.Text = "OEE";
                    //chart.Direction = eDirection.Column; // error: Property or indexer 'OfficeOpenXml.Drawing.Chart.ExcelBarChart.Direction' cannot be assigned to -- it is read only

                    ExcelRange erLossesRangeOEE = worksheet.Cells["M5:M" + (5 + TotalSummarizedRows)];
                    //chartOEE.YAxis.Fill.Color = System.Drawing.Color.Blue; //Working : 
                    //erLossesRangeOEE.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;//(Black)Color on DATA    //r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //erLossesRangeOEE.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                    //chartOEE.SetDataPointStyle(chartOEE, erLossesRangeOEE, Color.Red);
                    //chartOEE.Fill.Color = System.Drawing.Color.LightYellow; //Working :Charts Inside Background Color ~
                    //chartOEE.PlotArea.Fill.Color = System.Drawing.Color.LightYellow; //Charts Inside Background Color ~
                    chartOEE.Style = eChartStyle.None;
                    chartOEE.Legend.Remove();
                    chartOEE.DataLabel.ShowValue = true;
                    chartOEE.YAxis.MaxValue = 100;
                    chartOEE.YAxis.MinValue = 0;
                    chartOEE.XAxis.Font.Size = 8;
                    ////chartOEE.AxisX.IsMarginVisible = false;
                    ////chartOEE.AxisY.IsMarginVisible = false;
                    ////chartOEE.AxisX.LabelStyle.Format = "dd";
                    ////chartOEE.AxisX.MajorGrid.LineWidth = 0;
                    ////chartOEE.AxisY.MajorGrid.LineWidth = 0;
                    ////chartOEE.AxisY.LabelStyle.Font = new Font("Consolas", 8);
                    ////chartOEE.AxisX.LabelStyle.Font = new Font("Consolas", 8);
                    //chartOEE.ChartAreas[0].AxisX.LineDashStyle.Dot;
                    //chartOEE.ChartXml["barChartOEE"].x.MajorGrid.Enabled = false;
                    //Chart1.ChartAreas["ChartArea1"].AxisY.MajorGrid.Enabled = false;
                    chartOEE.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartOEE.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartOEE.Series.Add(erLossesRangeOEE, erLossesRangeMacInv);
                    //chartOEE.Legend.Border.Fill.Color = Color.Yellow;
                    RemoveGridLines(ref chartOEE1);

                    ////Get the nodes
                    //// const string PIE_PATH = "c:chartSpace/c:chart/c:plotArea/c:pie3DChart/c:ser";
                    //const string PIE_PATH = "c:chartSpace/c:chart/c:plotArea/c:barChart/c:dLbls";
                    //var ws = chartOEE.WorkSheet;
                    //var nsm = ws.Drawings.NameSpaceManager;
                    //var nschart = nsm.LookupNamespace("c");
                    //var nsa = nsm.LookupNamespace("a");
                    //var node = chartOEE.ChartXml.SelectSingleNode(PIE_PATH, nsm);
                    //var doc = chartOEE.ChartXml;

                    ////Add the node
                    //var rand = new Random();
                    //for (var i = 0; i < 4; i++)
                    //{
                    //    //Create the data point node
                    //    var dPt = doc.CreateElement("dPt", nschart);

                    //    var idx = dPt.AppendChild(doc.CreateElement("idx", nschart));
                    //    var valattrib = idx.Attributes.Append(doc.CreateAttribute("val"));
                    //    valattrib.Value = i.ToString(CultureInfo.InvariantCulture);
                    //    node.AppendChild(dPt);

                    //    //Add the solid fill node
                    //    var spPr = doc.CreateElement("spPr", nschart);
                    //    var solidFill = spPr.AppendChild(doc.CreateElement("solidFill", nsa));
                    //    var srgbClr = solidFill.AppendChild(doc.CreateElement("srgbClr", nsa));
                    //    valattrib = srgbClr.Attributes.Append(doc.CreateAttribute("val"));

                    //    //Set the color
                    //    var color = Color.FromArgb(rand.Next(256), rand.Next(256), rand.Next(256));
                    //    valattrib.Value = ColorTranslator.ToHtml(color).Replace("#", String.Empty);
                    //    dPt.AppendChild(spPr);
                    //}

                    #region Adds complete Grid Lines Not Using
                    ////Get reference to the worksheet xml for proper namespace
                    //var chartXml = chartOEE.ChartXml;
                    //var nsuri = chartXml.DocumentElement.NamespaceURI;
                    //var nsm = new XmlNamespaceManager(chartXml.NameTable);
                    //nsm.AddNamespace("c", nsuri);

                    ////XY Scatter plots have 2 value axis and no category
                    //var valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
                    //if (valAxisNodes != null && valAxisNodes.Count > 0)
                    //    foreach (XmlNode valAxisNode in valAxisNodes)
                    //    {
                    //        if (valAxisNode.SelectSingleNode("c:majorGridlines", nsm) == null)
                    //            valAxisNode.AppendChild(chartXml.CreateNode(XmlNodeType.Element, "c:majorGridlines", nsuri));
                    //        if (valAxisNode.SelectSingleNode("c:minorGridlines", nsm) == null)
                    //            valAxisNode.AppendChild(chartXml.CreateNode(XmlNodeType.Element, "c:minorGridlines", nsuri));
                    //    }

                    ////Other charts can have a category axis
                    //var catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
                    //if (catAxisNodes != null && catAxisNodes.Count > 0)
                    //    foreach (XmlNode catAxisNode in catAxisNodes)
                    //    {
                    //        if (catAxisNode.SelectSingleNode("c:majorGridlines", nsm) == null)
                    //            catAxisNode.AppendChild(chartXml.CreateNode(XmlNodeType.Element, "c:majorGridlines", nsuri));
                    //        if (catAxisNode.SelectSingleNode("c:minorGridlines", nsm) == null)
                    //            catAxisNode.AppendChild(chartXml.CreateNode(XmlNodeType.Element, "c:minorGridlines", nsuri));
                    //    }

                    #endregion

                    //Availability
                    ExcelChart chartAvail1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartAvail", eChartType.ColumnClustered);
                    var chartAvail = (ExcelBarChart)chartAvail1;
                    chartAvail.SetSize(300, 390);
                    chartAvail.SetPosition(60, 320);
                    ExcelRange erLossesRangechartAvail = worksheet.Cells["J5:J" + (5 + TotalSummarizedRows)];
                    //erLossesRangechartAvail.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //chartAvail.YAxis.Fill.Color = System.Drawing.Color.Orange;
                    chartAvail.Title.Text = "Availability  ";
                    chartAvail.Style = eChartStyle.Style12;
                    chartAvail.Legend.Remove();
                    chartAvail.YAxis.MaxValue = 100;
                    chartAvail.YAxis.MinValue = 0;
                    chartAvail.XAxis.Font.Size = 8;
                    chartAvail.DataLabel.ShowValue = true;
                    chartAvail.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartAvail.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartAvail.Series.Add(erLossesRangechartAvail, erLossesRangeMacInv);
                    RemoveGridLines(ref chartAvail1);

                    //Performance
                    ExcelChart chartPerf1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartPerf", eChartType.ColumnClustered);
                    var chartPerf = (ExcelBarChart)chartPerf1;
                    chartPerf.SetSize(300, 390);
                    chartPerf.SetPosition(60, 630);
                    chartPerf.XAxis.Font.Size = 8;
                    chartPerf.Title.Text = "Performance ";
                    ExcelRange erLossesRangechartPerf = worksheet.Cells["K5:K" + (5 + TotalSummarizedRows)];
                    //erLossesRangechartPerf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //chartPerf.YAxis.Fill.Color = System.Drawing.Color.Yellow;
                    chartPerf.Style = eChartStyle.Style11;
                    chartPerf.Legend.Remove();
                    chartPerf.YAxis.MaxValue = 100;
                    chartPerf.YAxis.MinValue = 0;
                    chartPerf.DataLabel.ShowValue = true;
                    chartPerf.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartPerf.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartPerf.Series.Add(erLossesRangechartPerf, erLossesRangeMacInv);
                    RemoveGridLines(ref chartPerf1);

                    //Quality
                    ExcelChart chartQual1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartQual", eChartType.ColumnClustered);
                    var chartQual = (ExcelBarChart)chartQual1;
                    chartQual.SetSize(300, 390);
                    chartQual.SetPosition(60, 940);
                    ExcelRange erLossesRangechartQual = worksheet.Cells["L5:L" + (5 + TotalSummarizedRows)];
                    chartQual.Title.Text = "Quality  ";
                    chartQual.Style = eChartStyle.Style9;
                    chartQual.Legend.Remove();
                    chartQual.YAxis.MaxValue = 100;
                    chartQual.YAxis.MinValue = 0;
                    chartQual.XAxis.Font.Size = 8;
                    chartQual.DataLabel.ShowValue = true;
                    chartQual.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartQual.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartQual.Series.Add(erLossesRangechartQual, erLossesRangeMacInv);
                    RemoveGridLines(ref chartQual1);

                    #endregion

                    #region Trend of OEE Over Selected DateRange
                    List<double> ForAvg = new List<double>();
                    UsedDateForExcel = Convert.ToDateTime(frmDate);
                    string CellsOfOEEYAxis = null;
                    string CellsOfOEEXAxis = null;
                    for (int i = 0; i < TotalDay + 1; i++)
                    {
                        string CorrectedDateString = UsedDateForExcel.ToString("yyyy-MM-dd");
                        var drs = from r in DTConsolidatedLosses.AsEnumerable()
                                  where (r.Field<string>("CorrectedDate") == CorrectedDateString && r.Field<string>("Plant") == "Summarized")
                                  select r;
                        int skipFirstRow = 0;
                        //if (drs != null)
                        //{
                        foreach (var cell in drs)
                        {
                            string CellRowString = cell["OEE"].ToString();
                            string CellRowDateString = cell["CorrectedDate"].ToString();

                            //Regex.Replace(str, "[^0-9 _]", "");
                            string extractedRowString = Regex.Replace(CellRowString, "[0-9 _]", string.Empty);
                            string extractedRowNumber = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                            if (CellsOfOEEXAxis == null)
                            {
                                CellsOfOEEXAxis = "H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }
                            else
                            {
                                CellsOfOEEXAxis += ",H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }

                            if (CellsOfOEEYAxis == null)
                            {
                                CellsOfOEEYAxis = CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }
                            else
                            {
                                CellsOfOEEYAxis += "," + CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }
                        }
                        //}
                        //else
                        //{}
                        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                    }


                    ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfOEEYAxis];
                    ExcelRange erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfOEEXAxis];
                    ExcelChart chartOEETrend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartOEE", eChartType.LineMarkers);

                    #region Experiment for Hybrid Graph Success
                    ////Now for the second chart type we use the chart.PlotArea.ChartTypes collection...
                    //worksheetGraph.Cells["B9:B10"].Value = "Avg";
                    //worksheetGraph.Cells["A9:A10"].Value = 30;
                    //var chartType2 = chartOEETrend1.PlotArea.ChartTypes.Add(eChartType.Line);
                    //var serie2 = chartType2.Series.Add(worksheetGraph.Cells["A9:A10"], worksheetGraph.Cells["B9:B10"]);

                    #endregion

                    for (int i = 0; i < ForAvg.Count; i++)
                    {
                        worksheetGraph.Cells["A" + (12 + i)].Value = "Avg";
                        worksheetGraph.Cells["B" + (12 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                    }

                    ////Based on TotalDays :: insert AvgValue into Cells
                    // var AvgVal = worksheetGraph.Calculate(" =AVERAGE(erLossesRangechartTop5LossesvALUE)");
                    // var AvgVal1 = worksheetGraph.Calculate(" =AVERAGE('" + erLossesRangechartTop5LossesvALUE + "')");
                    //worksheetGraph.Cells["F9"].Formula = " =AVERAGE(" + erLossesRangechartTop5LossesvALUE + ")";
                    //worksheetGraph.Cells["B9:B" + (9 + ForAvg.Count)].Value = "Avg";
                    //worksheetGraph.Cells["A9:A" + (9 + OEEsForAvg.Count)].Value = AvgVal;
                    //string blah = AvgVal.ToString();

                    //worksheetGraph.Cells["C9:C" + (9 + TotalDay + 1)].Value = AvgVal;
                    var chartType2 = chartOEETrend1.PlotArea.ChartTypes.Add(eChartType.Line);
                    var serie2 = chartType2.Series.Add(worksheetGraph.Cells["B12:B" + (12 + ForAvg.Count - 1)], worksheetGraph.Cells["B12:B" + (12 + ForAvg.Count - 1)]);

                    var chartOEETrend = (ExcelLineChart)chartOEETrend1;
                    chartOEETrend.SetSize(300, 300);
                    chartOEETrend.SetPosition(450, 10);
                    chartOEETrend.Title.Text = "OEE ";
                    chartOEETrend.Style = eChartStyle.Style4;
                    chartOEETrend.Legend.Remove();
                    chartOEETrend.YAxis.MaxValue = 100;
                    chartOEETrend.DataLabel.ShowValue = true;
                    chartOEETrend.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartOEETrend.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartOEETrend.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                    RemoveGridLines(ref chartOEETrend1);
                    #endregion

                    #region Trend of Availability Over Selected DateRange
                    UsedDateForExcel = Convert.ToDateTime(frmDate);
                    string CellsOfAvailYAxis = null;
                    string CellsOfAvailXAxis = null;
                    //ForAvg.RemoveAll(m=>m > -1);
                    ForAvg.Clear();
                    for (int i = 0; i < TotalDay + 1; i++)
                    {
                        string CorrectedDateString = UsedDateForExcel.ToString("yyyy-MM-dd");
                        var drs = from r in DTConsolidatedLosses.AsEnumerable()
                                  where (r.Field<string>("CorrectedDate") == CorrectedDateString && r.Field<string>("Plant") == "Summarized")
                                  select r;
                        int skipFirstRow = 0;
                        //if (drs != null)
                        //{
                        foreach (var cell in drs)
                        {
                            string CellRowString = cell["Avail"].ToString();
                            string CellRowDateString = cell["CorrectedDate"].ToString();

                            //Regex.Replace(str, "[^0-9 _]", "");
                            string extractedRowString = Regex.Replace(CellRowString, "[0-9 _]", string.Empty);
                            string extractedRowNumber = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                            if (CellsOfAvailXAxis == null)
                            {
                                CellsOfAvailXAxis = "H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }
                            else
                            {
                                CellsOfAvailXAxis += ",H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }

                            if (CellsOfAvailYAxis == null)
                            {
                                CellsOfAvailYAxis = CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }
                            else
                            {
                                CellsOfAvailYAxis += "," + CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }

                        }
                        //}
                        //else
                        //{}
                        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                    }

                    for (int i = 0; i < ForAvg.Count; i++)
                    {
                        worksheetGraph.Cells["C" + (12 + i)].Value = "Avg";
                        worksheetGraph.Cells["D" + (12 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                    }

                    erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfAvailYAxis];
                    erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfAvailXAxis];
                    ExcelChart chartAvailTrend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartAvail", eChartType.LineMarkers);

                    var chartAvailType2 = chartAvailTrend1.PlotArea.ChartTypes.Add(eChartType.Line);
                    var Availserie2 = chartAvailType2.Series.Add(worksheetGraph.Cells["D12:D" + (12 + ForAvg.Count - 1)], worksheetGraph.Cells["C12:C" + (12 + ForAvg.Count - 1)]);

                    var chartAvailTrend = (ExcelLineChart)chartAvailTrend1;
                    chartAvailTrend.SetSize(300, 300);
                    chartAvailTrend.SetPosition(450, 330);
                    chartAvailTrend.Title.Text = "Availability ";
                    chartAvailTrend.Style = eChartStyle.Style3;
                    chartAvailTrend.Legend.Remove();
                    chartAvailTrend.YAxis.MaxValue = 100;
                    chartAvailTrend.DataLabel.ShowValue = true;
                    chartAvailTrend.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartAvailTrend.XAxis.MajorTickMark = eAxisTickMark.None;
                    //chartAvailTrend.DataLabel.Position = eLabelPosition.InBase;
                    chartAvailTrend.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                    RemoveGridLines(ref chartAvailTrend1);
                    #endregion

                    #region Trend of Performance Over Selected DateRange
                    UsedDateForExcel = Convert.ToDateTime(frmDate);
                    string CellsOfPerfYAxis = null;
                    string CellsOfPerfXAxis = null;
                    ForAvg.Clear();
                    for (int i = 0; i < TotalDay + 1; i++)
                    {
                        string CorrectedDateString = UsedDateForExcel.ToString("yyyy-MM-dd");
                        var drs = from r in DTConsolidatedLosses.AsEnumerable()
                                  where (r.Field<string>("CorrectedDate") == CorrectedDateString && r.Field<string>("Plant") == "Summarized")
                                  select r;
                        int skipFirstRow = 0;
                        //if (drs != null)
                        //{
                        foreach (var cell in drs)
                        {
                            string CellRowString = cell["Perf"].ToString();
                            string CellRowDateString = cell["CorrectedDate"].ToString();

                            //Regex.Replace(str, "[^0-9 _]", "");
                            string extractedRowString = Regex.Replace(CellRowString, "[0-9 _]", string.Empty);
                            string extractedRowNumber = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                            if (CellsOfPerfXAxis == null)
                            {
                                CellsOfPerfXAxis = "H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }
                            else
                            {
                                CellsOfPerfXAxis += ",H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }

                            if (CellsOfPerfYAxis == null)
                            {
                                CellsOfPerfYAxis = CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }
                            else
                            {
                                CellsOfPerfYAxis += "," + CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }

                        }
                        //}
                        //else
                        //{}
                        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                    }

                    for (int i = 0; i < ForAvg.Count; i++)
                    {
                        worksheetGraph.Cells["E" + (12 + i)].Value = "Avg";
                        worksheetGraph.Cells["F" + (12 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                    }

                    erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfPerfYAxis];
                    erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfPerfXAxis];
                    ExcelChart chartPerfTrend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartPerf", eChartType.LineMarkers);

                    var chartPerfType2 = chartPerfTrend1.PlotArea.ChartTypes.Add(eChartType.Line);
                    var Perfserie2 = chartPerfType2.Series.Add(worksheetGraph.Cells["F12:F" + (12 + ForAvg.Count - 1)], worksheetGraph.Cells["E12:E" + (12 + ForAvg.Count - 1)]);

                    var chartPerfTrend = (ExcelLineChart)chartPerfTrend1;
                    //if (TotalDay < 7)
                    //{
                    chartPerfTrend.SetSize(300, 300);
                    chartPerfTrend.SetPosition(450, 640);
                    //}

                    chartPerfTrend.Title.Text = "Performance ";
                    chartPerfTrend.Style = eChartStyle.Style2;
                    chartPerfTrend.Legend.Remove();
                    chartPerfTrend.DataLabel.ShowValue = true;
                    chartPerfTrend.YAxis.MaxValue = 100;
                    chartPerfTrend.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartPerfTrend.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartPerfTrend.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                    RemoveGridLines(ref chartPerfTrend1);

                    #endregion

                    #region Trend of Quality Over Selected DateRange
                    UsedDateForExcel = Convert.ToDateTime(frmDate);
                    string CellsOfQualYAxis = null;
                    string CellsOfQualXAxis = null;
                    ForAvg.Clear();
                    for (int i = 0; i < TotalDay + 1; i++)
                    {
                        string CorrectedDateString = UsedDateForExcel.ToString("yyyy-MM-dd");
                        var drs = from r in DTConsolidatedLosses.AsEnumerable()
                                  where (r.Field<string>("CorrectedDate") == CorrectedDateString && r.Field<string>("Plant") == "Summarized")
                                  select r;
                        int skipFirstRow = 0;
                        //if (drs != null)
                        //{
                        foreach (var cell in drs)
                        {
                            string CellRowString = cell["Qual"].ToString();
                            string CellRowDateString = cell["CorrectedDate"].ToString();

                            //Regex.Replace(str, "[^0-9 _]", "");
                            string extractedRowString = Regex.Replace(CellRowString, "[0-9 _]", string.Empty);
                            string extractedRowNumber = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                            if (CellsOfQualXAxis == null)
                            {
                                CellsOfQualXAxis = "H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }
                            else
                            {
                                CellsOfQualXAxis += ",H" + (Convert.ToInt32(extractedRowNumber) - 1);
                            }

                            if (CellsOfQualYAxis == null)
                            {
                                CellsOfQualYAxis = CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }
                            else
                            {
                                CellsOfQualYAxis += "," + CellRowString;
                                double CellVal = 0;
                                if (double.TryParse(Convert.ToString(worksheet.Cells[CellRowString].Value), out CellVal))
                                {
                                    ForAvg.Add(CellVal);
                                }
                            }

                        }
                        //}
                        //else
                        //{}
                        UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                    }

                    for (int i = 0; i < ForAvg.Count; i++)
                    {
                        worksheetGraph.Cells["G" + (12 + i)].Value = "Avg";
                        worksheetGraph.Cells["H" + (12 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                    }


                    erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfQualYAxis];
                    erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfQualXAxis];
                    ExcelChart chartQualTrend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartQual", eChartType.LineMarkers);

                    var chartQualType2 = chartQualTrend1.PlotArea.ChartTypes.Add(eChartType.Line);
                    var Qualserie2 = chartQualType2.Series.Add(worksheetGraph.Cells["H12:H" + (12 + ForAvg.Count - 1)], worksheetGraph.Cells["G12:G" + (12 + ForAvg.Count - 1)]);
                    // var Perfserie2 = chartPerfType2.Series.Add(worksheetGraph.Cells["F12:F" + (12 + ForAvg.Count - 1)], worksheetGraph.Cells["E12:E" + (12 + ForAvg.Count - 1)]);

                    var chartQualTrend = (ExcelLineChart)chartQualTrend1;
                    chartQualTrend.SetSize(300, 300);
                    chartQualTrend.SetPosition(450, 950);
                    chartQualTrend.Title.Text = "Quality ";
                    chartQualTrend.Style = eChartStyle.Style1;
                    chartQualTrend.Legend.Remove();
                    chartQualTrend.YAxis.MaxValue = 100;
                    chartQualTrend.YAxis.MinValue = 0;
                    chartQualTrend.ShowDataLabelsOverMaximum = true;
                    chartQualTrend.DataLabel.ShowValue = true;
                    chartQualTrend.DataLabel.ShowBubbleSize = true;
                    chartQualTrend.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartQualTrend.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartQualTrend.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                    RemoveGridLines(ref chartQualTrend1);

                    #endregion

                    #region lOSSES TOP 5 GRAPH
                    //1. Get Top 5 Losses ColName in excel.
                    //2. Generate the comma seperated String format .
                    //sort the list
                    AllOccuredLosses.Sort(Compare2);
                    foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                    {
                        IntoFile(loss.Key);
                    }
                    AllOccuredLosses = AllOccuredLosses.GroupBy(x => x.Key).Select(g => g.First()).ToList();
                    AllOccuredLosses = AllOccuredLosses.OrderByDescending(x => x.Value).ToList();
                   

                    //Now construct string from top 5 losses.
                    int LooperTop5 = 0;
                    string CellsOfTop5LossColNames = null;
                    string CellsOfTop5LossColValues = null;
                    foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                    {
                        if (LooperTop5 < 5)
                        {
                            string a = loss.Key;
                            double b = loss.Value;

                            var outputJustColName = Regex.Replace(a, @"[\d-]", string.Empty);
                            //string LossCol = Convert.ToString(outputJustColName);
                            string LossCol = a;
                            if (LooperTop5 == 0)
                            {
                                CellsOfTop5LossColNames = LossCol;
                                CellsOfTop5LossColValues = outputJustColName + Row;
                            }
                            else
                            {
                                CellsOfTop5LossColNames += "," + LossCol;
                                CellsOfTop5LossColValues += "," + outputJustColName + Row;
                            }
                        }
                        else
                        {
                            break;
                        }
                        LooperTop5++;
                    }

                    //IntoFile(CellsOfTop5LossColNames + "\n");
                    //IntoFile(CellsOfTop5LossColValues + " ");

                    //To make sure it doesn't through error when there's no data.
                    if (AllOccuredLosses.Count > 0)
                    {
                        ExcelChart chartTop5Losses1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartTop5Losses", eChartType.ColumnClustered);
                        var chartTop5Losses = (ExcelBarChart)chartTop5Losses1;
                        chartTop5Losses.SetSize(500, 400);
                        chartTop5Losses.SetPosition(760, 10);
                        //string blah = "CY11,CZ11,DA11,DC11,DD11"; //This Works 
                        //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells[blah];
                        //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells["CY11,CZ11,DA11,DC11,DD11"];
                        //ExcelRange erLossesRangechartTop5LossesNAMES = worksheet.Cells["CY3,CZ3,DA3,DC3,DD3"];

                        //Assign Null
                        erLossesRangechartTop5LossesvALUE = null;
                        erLossesRangechartTop5LossesNAMES = null;
                        ExcelRange erLossesRangechartTop5LossesvALUE6 = worksheet.Cells[CellsOfTop5LossColValues];
                        ExcelRange erLossesRangechartTop5LossesNAMES6 = worksheet.Cells[CellsOfTop5LossColNames];
                        chartTop5Losses.Title.Text = "LOSSES (Hrs)";
                        chartTop5Losses.Style = eChartStyle.Style19;
                        chartTop5Losses.Legend.Remove();
                        chartTop5Losses.DataLabel.ShowValue = true;
                        //chartTop5Losses.DataLabel.Font.Size = 8;
                        //chartTop5Losses.Legend.Font.Size = 8;
                        chartTop5Losses.XAxis.Font.Size = 8;
                        chartTop5Losses.YAxis.MinorTickMark = eAxisTickMark.None;
                        chartTop5Losses.XAxis.MajorTickMark = eAxisTickMark.None;
                        //IntoFile(erLossesRangechartTop5LossesvALUE6 + "\n");
                        //IntoFile(erLossesRangechartTop5LossesNAMES6 + " ");
                        //chartTop5Losses.Series.Chart.XAxis.d
                        chartTop5Losses.Series.Add(erLossesRangechartTop5LossesvALUE6, erLossesRangechartTop5LossesNAMES6);
                        RemoveGridLines(ref chartTop5Losses1);

                    }
                    #endregion

                    #region Identified & UnIdentified Losses "scary"
                    worksheetGraph.Cells["A1"].Value = "Ratio of Losses";
                    //worksheetGraph.Cells["A2"].Value = "UnIdentifiedLoss";

                    double IdentifiedLossPercentage = (IdentifiedLoss / (IdentifiedLoss + UnIdentifiedLoss)) * 100;
                    double UnIdentifiedLossPercentage = (UnIdentifiedLoss / (IdentifiedLoss + UnIdentifiedLoss)) * 100;
                    worksheetGraph.Cells["B1"].Value = Math.Round(IdentifiedLossPercentage, 0);
                    worksheetGraph.Cells["B2"].Value = Math.Round(UnIdentifiedLossPercentage, 0);

                    erLossesRangechartTop5LossesvALUE = worksheetGraph.Cells["B1"];
                    ExcelRange erLossesRangechartTop5LossesvALUE1 = worksheetGraph.Cells["B2"];
                    erLossesRangechartTop5LossesNAMES = worksheetGraph.Cells["A1"];

                    ExcelChart chartIDAndUnID1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("TypesOfLosses", eChartType.ColumnStacked);
                    var chartIDAndUnID = (ExcelBarChart)chartIDAndUnID1;
                    chartIDAndUnID.SetSize(500, 350);
                    chartIDAndUnID.SetPosition(760, 520);

                    chartIDAndUnID.Title.Text = "Identified Losses  ";
                    chartIDAndUnID.Style = eChartStyle.Style18;
                    chartIDAndUnID.Legend.Position = eLegendPosition.Bottom;
                    //chartIDAndUnID.Legend.Remove();
                    chartIDAndUnID.YAxis.MaxValue = 100;
                    chartIDAndUnID.YAxis.MinValue = 0;
                    chartIDAndUnID.Locked = false;
                    chartIDAndUnID.PlotArea.Border.Width = 0;
                    chartIDAndUnID.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartIDAndUnID.DataLabel.ShowValue = true;
                    //chartAllLosses.DataLabel.ShowValue = true;
                    var thisYearSeries = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES));
                    thisYearSeries.Header = "Identified Losses";
                    var lastYearSeries = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erLossesRangechartTop5LossesvALUE1, erLossesRangechartTop5LossesNAMES));
                    lastYearSeries.Header = "UnIdentified Losses";
                    RemoveGridLines(ref chartIDAndUnID1);

                    #region OLD
                    ////////////////////
                    //have to remove cat nodes from each series so excel autonums 1 and 2 in xaxis
                    //var chartXml = chartIDAndUnID.ChartXml;
                    //var nsm = new XmlNamespaceManager(chartXml.NameTable);

                    //var nsuri = chartXml.DocumentElement.NamespaceURI;
                    //nsm.AddNamespace("c", nsuri);

                    ////Get the Series ref and its cat
                    //var serNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser", nsm);
                    //foreach (XmlNode serNode in serNodes)
                    //{
                    //    //Cell any cell reference and replace it with a string literal list
                    //    var catNode = serNode.SelectSingleNode("c:cat", nsm);
                    //    catNode.RemoveAll();

                    //    //Create the string list elements
                    //    var ptCountNode = chartXml.CreateElement("c:ptCount", nsuri);
                    //    ptCountNode.Attributes.Append(chartXml.CreateAttribute("val", nsuri));
                    //    ptCountNode.Attributes[0].Value = "2";

                    //    var v0Node = chartXml.CreateElement("c:v", nsuri);
                    //    v0Node.InnerText = "opening";
                    //    var pt0Node = chartXml.CreateElement("c:pt", nsuri);
                    //    pt0Node.AppendChild(v0Node);
                    //    pt0Node.Attributes.Append(chartXml.CreateAttribute("idx", nsuri));
                    //    pt0Node.Attributes[0].Value = "0";

                    //    var v1Node = chartXml.CreateElement("c:v", nsuri);
                    //    v1Node.InnerText = "closing";
                    //    var pt1Node = chartXml.CreateElement("c:pt", nsuri);
                    //    pt1Node.AppendChild(v1Node);
                    //    pt1Node.Attributes.Append(chartXml.CreateAttribute("idx", nsuri));
                    //    pt1Node.Attributes[0].Value = "1";

                    //    //Create the string list node
                    //    var strLitNode = chartXml.CreateElement("c:strLit", nsuri);
                    //    strLitNode.AppendChild(ptCountNode);
                    //    strLitNode.AppendChild(pt0Node);
                    //    strLitNode.AppendChild(pt1Node);
                    //    catNode.AppendChild(strLitNode);
                    //}
                    //pck.Save();
                    #endregion
                    #region Experiment to Send Data to Excel Chart in Template
                    //OfficeOpenXml.FormulaParsing.Excel.Application xlApp;
                    //Excel.Workbook xlWorkBook;
                    //Excel.Worksheet xlWorkSheet;
                    //object misValue = System.Reflection.Missing.Value;

                    //Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                    //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                    //Excel.Chart chartPage = myChart.Chart;

                    //Excel.Range chartRange;
                    //chartRange = xlWorkSheet.get_Range("A1", "d5");
                    //chartPage.SetSourceData(chartRange, misValue);
                    //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

                    //xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    //xlWorkBook.Close(true, misValue, misValue);
                    //xlApp.Quit();

                    //var piechart = worksheet.Drawings["A"] as ExcelBarChart;
                    //piechart.Style = eChartStyle.Style26;
                    //chartIDAndUnID.Style = eChartStyle.Style26;

                    //////////////////////////////


                    #endregion

                    #endregion End of Identified & UnIdentified Losses

                    #region All Losses Chart

                    CellsOfTop5LossColNames = null;
                    CellsOfTop5LossColValues = null;
                    int Looper = 0;
                    foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                    {
                        string a = loss.Key;
                        double b = loss.Value;

                        var outputJustColName = Regex.Replace(a, @"[\d-]", string.Empty);
                        //string LossCol = Convert.ToString(outputJustColName);
                        string LossCol = a;
                        if (a == "0")
                        {
                        }
                        if (Looper == 0)
                        {
                            CellsOfTop5LossColNames = LossCol;
                            CellsOfTop5LossColValues = outputJustColName + Row;
                        }
                        else
                        {
                            CellsOfTop5LossColNames += "," + LossCol;
                            CellsOfTop5LossColValues += "," + outputJustColName + Row;
                        }
                        Looper++;
                    }

                    if (AllOccuredLosses.Count > 0)
                    {
                        ExcelChart chartAllLosses1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartAllLosses", eChartType.ColumnClustered);
                        var chartAllLosses = (ExcelBarChart)chartAllLosses1;
                        chartAllLosses.SetSize(1200, 500);
                        chartAllLosses.SetPosition(1170, 10);
                        erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfTop5LossColValues];
                        erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfTop5LossColNames];
                        chartAllLosses.Title.Text = "All LOSSES" + "(Hrs)";
                        chartAllLosses.Style = eChartStyle.Style25;
                        chartAllLosses.Legend.Remove();
                        chartAllLosses.DataLabel.ShowValue = true;
                        //chartAllLosses.DataLabel.Font.Size = 8;
                        //chartAllLosses.Legend.Font.Size = 8;
                        chartAllLosses.YAxis.MinorTickMark = eAxisTickMark.None;
                        chartAllLosses.XAxis.MajorTickMark = eAxisTickMark.None;
                        chartAllLosses.XAxis.Font.Size = 8;
                        chartAllLosses.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);

                        //Get reference of Graph to Remove GridLines
                        RemoveGridLines(ref chartAllLosses1);

                    }

                    #endregion

                    #region Experiment Linq on Excel . IT's a Success

                    //foreach (var MacINV in WCInvNoList)
                    //{
                    //    string WCInvNoString = Convert.ToString(MacINV);

                    //    //Select all cells in column G with this MacINV
                    //    var queryLinq = (from cell in worksheet.Cells["G:G"]
                    //                     where cell.Value is string && (string)cell.Value == @WCInvNoString 
                    //                     select cell);
                    //    foreach (var cell in queryLinq)
                    //    {
                    //        string CellRowString = cell.Address;
                    //    }

                    //}
                    #endregion

                    #region  Losses Trend :: All 5 Topper's
                    var queryLinq = (from cell in worksheet.Cells["C:C"]
                                     where cell.Value is string && (string)cell.Value == "Summarized For"
                                     select cell);
                    int LossesLooper = 1;
                    int PositionY = 1690;
                    int PositionX = 10;

                    int GraphNo = 0;
                    foreach (var Loss in AllOccuredLosses)
                    {
                        if (LossesLooper <= 5)
                        {
                            ForAvg.Clear();
                            CellsOfOEEYAxis = null;
                            CellsOfOEEXAxis = null;
                            string CellColString = Convert.ToString(Loss.Key);
                            string LossName = Convert.ToString(worksheet.Cells[CellColString].Value);
                            foreach (var cell in queryLinq)
                            {
                                string CellRowString = cell.Address;
                                string RowNum = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                                string ColName = Regex.Replace(CellColString, "[0-9 _]", string.Empty);

                                if (CellsOfOEEXAxis == null)
                                {
                                    CellsOfOEEXAxis = "H" + (Convert.ToInt32(RowNum) - 1);
                                }
                                else
                                {
                                    CellsOfOEEXAxis += ",H" + (Convert.ToInt32(RowNum) - 1);
                                }

                                if (CellsOfOEEYAxis == null)
                                {
                                    string colPlusrow = ColName + RowNum;
                                    CellsOfOEEYAxis = colPlusrow;
                                    double CellVal = 0;
                                    string CellValString = worksheet.Calculate(worksheet.Cells[colPlusrow].Formula).ToString();
                                    if (double.TryParse(CellValString, out CellVal))
                                    {
                                        ForAvg.Add(CellVal);
                                    }
                                    //if (double.TryParse(Convert.ToString(worksheet.Cells[colPlusrow].Value), out CellVal))
                                    //{
                                    //    ForAvg.Add(CellVal);
                                    //}
                                }
                                else
                                {
                                    string colPlusrow = ColName + RowNum;
                                    CellsOfOEEYAxis += "," + colPlusrow;
                                    double CellVal = 0;

                                    string CellValString = worksheet.Calculate(worksheet.Cells[colPlusrow].Formula).ToString();
                                    if (double.TryParse(CellValString, out CellVal))
                                    {
                                        ForAvg.Add(CellVal);
                                    }
                                }
                            }

                            ExcelRange erTopLossesTrendYValues = worksheet.Cells[CellsOfOEEYAxis];
                            ExcelRange erTopLossesTrendXNames = worksheet.Cells[CellsOfOEEXAxis];
                            ExcelChart chartLossses1Trend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartTop" + LossesLooper, eChartType.LineMarkers);

                            GraphNo++;
                            if (GraphNo == 1)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["A" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["B" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["B50:B" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["A50:A" + (50 + ForAvg.Count - 1)]);
                            }
                            if (GraphNo == 2)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["C" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["D" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["D50:D" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["C50:C" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 3)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["E" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["F" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["F50:F" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["E50:E" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 4)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["G" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["H" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["H50:H" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["G50:G" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 5)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["I" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["J" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["J50:J" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["I50:I" + (50 + ForAvg.Count - 1)]);

                            }

                            var chartLossses1Trend = (ExcelLineChart)chartLossses1Trend1;
                            chartLossses1Trend.SetSize(300, 230);
                            chartLossses1Trend.SetPosition(PositionY, 10 + (((LossesLooper - 1)) * 300));
                            //chartLossses1Trend.Title.Text = "Top" + LossesLooper + " Trend Chart " + "(Hrs)";
                            chartLossses1Trend.Title.Text = LossName;
                            chartLossses1Trend.Style = eChartStyle.Style8;
                            chartLossses1Trend.Legend.Remove();
                            //chartLossses1Trend.YAxis.MaxValue = 100;
                            chartLossses1Trend.DataLabel.ShowValue = true;
                            //chartLossses1Trend.DataLabel.Font.Size = 6.0F;
                            chartLossses1Trend.PlotArea.Border.Width = 0;
                            chartLossses1Trend.YAxis.MinorTickMark = eAxisTickMark.None;
                            //chartLossses1Trend.YAxis.MajorTickMark = eAxisTickMark.None;
                            //chartLossses1Trend.XAxis.MinorTickMark = eAxisTickMark.None;
                            chartLossses1Trend.XAxis.MajorTickMark = eAxisTickMark.None;
                            //chartLossses1Trend.XAxis.MinorTickMark = eAxisTickMark.None;
                            chartLossses1Trend.Series.Add(erTopLossesTrendYValues, erTopLossesTrendXNames);

                            //Get reference of Graph to Remove GridLines
                            RemoveGridLines(ref chartLossses1Trend1);

                            #region OLD
                            //chartXml = chartLossses1Trend.ChartXml;
                            // nsuri = chartXml.DocumentElement.NamespaceURI;
                            // nsm = new XmlNamespaceManager(chartXml.NameTable);
                            //nsm.AddNamespace("c", nsuri);

                            ////XY Scatter plots have 2 value axis and no category
                            //valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
                            //if (valAxisNodes != null && valAxisNodes.Count > 0)
                            //    foreach (XmlNode valAxisNode in valAxisNodes)
                            //    {
                            //        var major = valAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                            //        if (major != null)
                            //            valAxisNode.RemoveChild(major);

                            //        var minor = valAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                            //        if (minor != null)
                            //            valAxisNode.RemoveChild(minor);
                            //    }

                            ////Other charts can have a category axis
                            //catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
                            //if (catAxisNodes != null && catAxisNodes.Count > 0)
                            //    foreach (XmlNode catAxisNode in catAxisNodes)
                            //    {
                            //        var major = catAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                            //        if (major != null)
                            //            catAxisNode.RemoveChild(major);

                            //        var minor = catAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                            //        if (minor != null)
                            //            catAxisNode.RemoveChild(minor);
                            //    }
                            #endregion

                            LossesLooper++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    #endregion
                }
                #endregion //End of Graphs

                //To Set Colors
                //http://stackoverflow.com/questions/36520427/legend-color-is-incorrect-in-excel-chart-created-using-epplus/36532733#36532733

                //Hide Column R(CuttingTime)
                worksheet.Column(18).Width = 0;
                worksheetGraph.Row(Row).Height = 0;

                //Hide Identified and UnIdentified Losses Values
                //Color ColorHexWhite = System.Drawing.Color.White;
                //worksheetGraph.Cells["A1:B2"].Style.Font.Color.SetColor(ColorHexWhite);
                //worksheetGraph.Cells["A3:Z100"].Style.Font.Color.SetColor(ColorHexWhite);

                //autofit
                //Apply style to Losses header
                Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#32CD32");//#32CD32:lightgreen //B8C9E9
                worksheet.Cells["W" + 3 + ":" + finalLossCol + "" + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["W" + 3 + ":" + finalLossCol + "" + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
                worksheet.Cells["W" + 3 + ":" + finalLossCol + "" + 4].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                worksheet.Cells["W" + 3 + ":" + finalLossCol + "" + 4].Style.WrapText = true;

                //For Header
                worksheet.Cells["F2:L2"].Merge = true;
                worksheet.Cells["F2:I2"].Style.Font.Bold = true;
                worksheet.Cells["F2:I2"].Style.Font.Size = 16;
                worksheet.Cells["F2"].Value = Header.ToUpper() + " OEE Analysis";
                worksheetGraph.Cells["F2:L2"].Merge = true;
                worksheetGraph.Cells["F2"].Style.Font.Bold = true;
                worksheetGraph.Cells["F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheetGraph.Cells["F2"].Style.Font.Size = 16;
                worksheetGraph.Cells["F2"].Value = Header.ToUpper() + " OEE Analysis";
                worksheetGraph.Cells["F2:M2"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                //worksheetGraph.Cells["A4:R100"].Style.Font.Color.SetColor(Color.White);

                worksheet.View.ShowGridLines = false;
                worksheetGraph.View.ShowGridLines = false;
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                p.Save();

                //Downloding Excel
                string path1 = System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
                if (TimeFactor == "GH")
                {
                    path1 = System.IO.Path.Combine(FileDir, "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"); //OEEReportAdjusted
                }
                else if (TimeFactor == "NoBlue")
                {
                    path1 = System.IO.Path.Combine(FileDir, "OEEReportAdjusted" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
                }
                System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
                string Outgoingfile = "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx";
                if (TimeFactor == "GH")
                {
                    Outgoingfile = "OEEReportGodHours" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx"; //OEEReportAdjusted
                }
                else if (TimeFactor == "NoBlue")
                {
                    Outgoingfile = "OEEReportAdjusted" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx";
                }
                if (file1.Exists)
                {
                    //Response.Clear();
                    //Response.ClearContent();
                    //Response.ClearHeaders();
                    //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                    //Response.AddHeader("Content-Length", file1.Length.ToString());
                    //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    //Response.WriteFile(file1.FullName);
                    //Response.Flush();
                    //Response.Close();
                }
                return path1;
            }
            catch (Exception e)
            {
               IntoFile("OEE Report: " + e);
            }
            return null;
        }

        //code to remove major GridLines
        public void RemoveGridLines(ref ExcelChart chartName)
        {
            var chartXml = chartName.ChartXml;
            var nsuri = chartXml.DocumentElement.NamespaceURI;
            var nsm = new XmlNamespaceManager(chartXml.NameTable);
            nsm.AddNamespace("c", nsuri);

            //XY Scatter plots have 2 value axis and no category
            var valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
            if (valAxisNodes != null && valAxisNodes.Count > 0)
                foreach (XmlNode valAxisNode in valAxisNodes)
                {
                    var major = valAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                    if (major != null)
                        valAxisNode.RemoveChild(major);

                    var minor = valAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                    if (minor != null)
                        valAxisNode.RemoveChild(minor);
                }

            //Other charts can have a category axis
            var catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
            if (catAxisNodes != null && catAxisNodes.Count > 0)
                foreach (XmlNode catAxisNode in catAxisNodes)
                {
                    var major = catAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                    if (major != null)
                        catAxisNode.RemoveChild(major);

                    var minor = catAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                    if (minor != null)
                        catAxisNode.RemoveChild(minor);
                }
        }

        static int Compare2(KeyValuePair<string, double> a, KeyValuePair<string, double> b)
        {
            return a.Value.CompareTo(b.Value);
        }
        //Sort KeyValuePair List Based on Key
        static int Compare1(KeyValuePair<string, double> a, KeyValuePair<string, double> b)
        {
            return a.Key.CompareTo(b.Key);
        }

        public async Task<string> ManualWOReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\ManualWOReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateSummarized = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                MacCount = 1;
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
            }

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "ManualWOReport" + lowestLevelName + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "ManualWOReport" + lowestLevelName + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    ////TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                    IntoFile("Excel with same date is already open, please close it and try to generate!!!!");
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetOA = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            } 
            if (worksheetOA == null)
            {
                try{
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //Step1: Get the Summarized according to Name. for each WC from StartDate - EndDate
            //( Leave that many rows blank and Fill WorkCenter wise 1st.
           

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //Excel Starts @ 5th row (i.e After Header) . => Row = 5 + MacCount + 1 ( Extra row after Overall Data ).
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5;
            int Sno = 1;
            for (int i = 0; i < TotalDay + 1; i++)
            {
                int thisDatesDataStartsAt = Row;
                string DayStartTime = Convert.ToDateTime(UsedDateForExcel).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);
                string DayEndTime = Convert.ToDateTime(UsedDateForExcel).AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);

                DateTime endDateTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = UsedDateForExcel.ToString("yyyy-MM-dd");

                DataTable HMIData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = null;
                    if (lowestLevel == "Plant")
                    {
                        //query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI where DDLWokrCentre is null and IsMultiWO = 0 and ( isWorkInProgress = 1 || isWorkInProgress = 0 ) and CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where PlantID = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) )) order by CorrectedDate ;";
                        query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI where DDLWokrCentre is null and IsMultiWO = 0 and ( isWorkInProgress = 1 or isWorkInProgress = 0 ) and CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where PlantID = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) )) order by CorrectedDate ;";
                    }
                    else if (lowestLevel == "Shop")
                    {
                        //query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 || isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) )) order by CorrectedDate ;";
                        query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 or isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) )) order by CorrectedDate ;";
                    }
                    else if (lowestLevel == "Cell")
                    {
                        //query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 || isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) )) order by CorrectedDate ;";
                        query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 or isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) )) order by CorrectedDate ;";
                    }
                    else if (lowestLevel == "WorkCentre")
                    {
                        //query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 || isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) )) order by CorrectedDate ;";
                        query = "select * from i_facility_tal.dbo.tblhmiscreen as tblHMI  where DDLWokrCentre is null  and IsMultiWO = 0  and ( isWorkInProgress = 1 or isWorkInProgress = 0 )  and  CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' and  MachineID in (select MachineID from [i_facility_tal].[dbo].tblmachinedetails where MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) )) order by CorrectedDate ;";
                    }

                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    //SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(HMIData);
                    mc.close();
                }

                for (int n = 0; n < HMIData.Rows.Count; n++)
                {
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        thisDatesDataStartsAt = Row;
                    }
                    int MachineID = Convert.ToInt32(HMIData.Rows[n][1]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = HierarchyData[3];

                    worksheet.Cells["G" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                    worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                    worksheet.Cells["I" + Row].Value = Convert.ToString(HMIData.Rows[n][3]);
                    worksheet.Cells["J" + Row].Value = Convert.ToString(HMIData.Rows[n][20]);

                    worksheet.Cells["K" + Row].Value = Convert.ToString(HMIData.Rows[n][6]);
                    worksheet.Cells["L" + Row].Value = Convert.ToString(HMIData.Rows[n][10]);
                    worksheet.Cells["M" + Row].Value = Convert.ToString(HMIData.Rows[n][7]);

                    worksheet.Cells["N" + Row].Value = Convert.ToString(HMIData.Rows[n][8]);
                    worksheet.Cells["O" + Row].Value = Convert.ToString(HMIData.Rows[n][15]);
                    worksheet.Cells["P" + Row].Value = Convert.ToString(HMIData.Rows[n][11]);
                    worksheet.Cells["Q" + Row].Value = Convert.ToString(HMIData.Rows[n][12]);
                    worksheet.Cells["R" + Row].Value = Convert.ToString(HMIData.Rows[n][4]);
                    worksheet.Cells["S" + Row].Value = Convert.ToString(HMIData.Rows[n][5]);

                    //double Minutes = Math.Round(Convert.ToDateTime(HMIData.Rows[n][5]).Subtract(Convert.ToDateTime(HMIData.Rows[n][4])).TotalMinutes, 2);
                    //worksheet.Cells["T" + Row].Value = Convert.ToString(Minutes);
                    Row++;
                }

                if (Row > thisDatesDataStartsAt)
                {
                    //thin border for all cells
                    var modelTable = worksheet.Cells["B" + thisDatesDataStartsAt + ":S" + (Row - 1)];
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + thisDatesDataStartsAt + ":S" + "" + (Row - 1)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                }

                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }

            worksheet.Cells["B3:S" + Row].AutoFitColumns();
            worksheet.Cells["B4:S4"].Style.WrapText = true;

            //To Push Summarized Data
            //get machines to be included in query
            DateTime endDateTime1 = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime1 = frmDate.ToString("yyyy-MM-dd");

            DataTable MacData = new DataTable();
            using (MsqlConnection mc1 = new MsqlConnection())
            {
                mc1.open();
                String queryMacs = null;
                if (lowestLevel == "Plant")
                {
                    //queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where PlantID = " + PlantID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) end) ); ";
                    queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where PlantID = " + PlantID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) ) ); ";
                }
                else if (lowestLevel == "Shop")
                {
                    //queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where ShopID = " + ShopID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) end) ); ";
                    queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where ShopID = " + ShopID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) ) ); ";
                }
                else if (lowestLevel == "Cell")
                {
                    //queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where CellID = " + CellID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) end) ); ";
                    queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where CellID = " + CellID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) ) ); ";
                }
                else if (lowestLevel == "WorkCentre")
                {
                    //queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) end) ); ";
                    queryMacs = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime1.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime1 + "' ) ) ); ";
                }

                //SqlDataAdapter da1 = new SqlDataAdapter(queryMacs, mc1.msqlConnection);
                SqlDataAdapter da1 = new SqlDataAdapter(queryMacs, mc1.msqlConnection);
                da1.Fill(MacData);
                mc1.close();
            }

            Row = 6;
            Sno = 1;
            int ManualWOCountSummarized = 0;
            int DDLWOCountSummarized = 0;
            for (int n = 0; n < MacData.Rows.Count; n++)
            {
                int MachineID = Convert.ToInt32(MacData.Rows[n][0]);
                List<string> HierarchyData = GetHierarchyData(MachineID);
                int ManualWOCount = 0, DDLWOCount = 0;
                string SDate = Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd");
                string DayStartTime = (Convert.ToDateTime(StartDate)).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);
                string DayEndTime = (Convert.ToDateTime(EndDate).AddDays(1)).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);

                DataTable MacDataCount = new DataTable();
                using (MsqlConnection mcCount = new MsqlConnection())
                {
                    mcCount.open();
                    String queryMacsCount = "select Count(*) from i_facility_tal.dbo.tblhmiscreen as tblHMI where DDLWokrCentre is null and IsMultiWO = 0  and MachineID = " + MachineID + " and  (isWorkInProgress = 0 or isWorkInProgress = 1) and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' ;";
                    //SqlDataAdapter daCount = new SqlDataAdapter(queryMacsCount, mcCount.msqlConnection);
                    SqlDataAdapter daCount = new SqlDataAdapter(queryMacsCount, mcCount.msqlConnection);
                    daCount.Fill(MacDataCount);
                    mcCount.close();
                }

                try
                {
                    if (MacDataCount.Rows.Count > 0)
                    {
                        string WOCountString = Convert.ToString(MacDataCount.Rows[0][0]);
                        int.TryParse(WOCountString, out ManualWOCount);
                        ManualWOCountSummarized += ManualWOCount;
                    }
                }
                catch (Exception e)
                {
                }

                DataTable MacDataCountDDL = new DataTable();
                using (MsqlConnection mcCountDDL = new MsqlConnection())
                {
                    mcCountDDL.open();
                    String queryMacsCountDDL = "select Count(*) from i_facility_tal.dbo.tblhmiscreen as tblHMI where ( DDLWokrCentre is not null or IsMultiWO = 1 )  and MachineID = " + MachineID + " and  (isWorkInProgress = 0 or isWorkInProgress = 1) and tblHMI.Time >= '" + DayStartTime + "'  and tblHMI.Time <= '" + DayEndTime + "' ;";
                    //SqlDataAdapter daCountDDL = new SqlDataAdapter(queryMacsCountDDL, mcCountDDL.msqlConnection);
                    SqlDataAdapter daCountDDL = new SqlDataAdapter(queryMacsCountDDL, mcCountDDL.msqlConnection);
                    daCountDDL.Fill(MacDataCountDDL);
                    mcCountDDL.close();
                }

                try
                {
                    if (MacDataCountDDL.Rows.Count > 0)
                    {
                        string DDLWOCountString = Convert.ToString(MacDataCountDDL.Rows[0][0]);
                        int.TryParse(DDLWOCountString, out DDLWOCount);
                        DDLWOCountSummarized += DDLWOCount;
                    }
                }
                catch (Exception e)
                {
                }

                worksheetOA.Cells["B" + Row].Value = Sno++;
                worksheetOA.Cells["C" + Row].Value = HierarchyData[0];
                worksheetOA.Cells["D" + Row].Value = HierarchyData[1];
                worksheetOA.Cells["E" + Row].Value = HierarchyData[2];
                worksheetOA.Cells["F" + Row].Value = HierarchyData[3];

                worksheetOA.Cells["G" + Row].Value = Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd");
                worksheetOA.Cells["H" + Row].Value = Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd");
                worksheetOA.Cells["I" + Row].Value = ManualWOCount;
                worksheetOA.Cells["J" + Row].Value = DDLWOCount;
                Row++;
            }

            //Row = 6;
            //thin border for all cells
            var modelTable1 = worksheetOA.Cells["B5:J" + Row];
            // Assign borders
            modelTable1.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            worksheetOA.Cells["I" + Row].Value = ManualWOCountSummarized;
            worksheetOA.Cells["J" + Row].Value = DDLWOCountSummarized;

            double PercentageManualWO = 0;
            double PercentageDDLWO = 0;
            if (ManualWOCountSummarized == 0 || DDLWOCountSummarized == 0)
            {
            }
            else
            {
                double tempSum = (ManualWOCountSummarized + DDLWOCountSummarized);
                double tempDiv = (ManualWOCountSummarized / tempSum);
                PercentageManualWO = (tempDiv * 100);

                double tempDiv1 = (DDLWOCountSummarized / tempSum);
                PercentageDDLWO = (tempDiv1 * 100);
            }

            worksheetOA.Cells["I" + (Row + 1)].Value = Math.Round(PercentageManualWO, 0) + "%";
            worksheetOA.Cells["J" + (Row + 1)].Value = Math.Round(PercentageDDLWO, 0) + "%";

            worksheetOA.Cells["I" + (Row + 1)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
            worksheetOA.Cells["J" + (Row + 1)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            worksheetOA.Cells["B" + Row + ":H" + Row].Merge = true;
            worksheetOA.Cells["B" + Row].Value = "Total Entry Events";
            worksheetOA.Cells["B" + Row + ":J" + (Row + 1)].Style.Font.Bold = true;

            //Excel:: Border Around Cells.
            worksheetOA.Cells["B5:J" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
            worksheet.View.ShowGridLines = false;
            worksheetOA.View.ShowGridLines = false;

            worksheetOA.Cells["B4:J" + (Row + 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //worksheet.View.FreezePanes(6, 1);
            // worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            //worksheetOA.View.FreezePanes(6, 1);
            //worksheetOA.Cells[worksheetOA.Dimension.Address].AutoFitColumns();

            worksheetOA.Cells["B4:H" + Row].AutoFitColumns();

            if (ManualWOCountSummarized == 0 && DDLWOCountSummarized == 0)
            { }
            else
            {
                //PIE Chart for OverAll Manual v/s DDL WO Count
                string ValCells = "I" + (Row) + ",J" + (Row);
                string NameCells = "I5,J5";
                ExcelRange erValues = worksheetOA.Cells[ValCells];
                ExcelRange erNames = worksheetOA.Cells[NameCells];
                var chartWOCount = (ExcelPieChart)worksheetOA.Drawings.AddChart("PieChartWOCount", eChartType.Pie);
                chartWOCount.SetSize(325, 325);
                chartWOCount.SetPosition(4, 0, 11, 0);
                //chartWOCount.SetPosition((Row+2),0,2,0);
                chartWOCount.Title.Text = "WO Entry Types";
                chartWOCount.Style = eChartStyle.Style18;
                chartWOCount.Series.Add(erValues, erNames);
                //chartWOCount.DataLabel.ShowSeriesName = true;
                chartWOCount.DataLabel.ShowCategory = true;
                chartWOCount.DataLabel.ShowPercent = true;
                //chartWOCount.DataLabel.ShowLeaderLines = true;
                chartWOCount.Legend.Remove();

                double YInPixels = worksheetOA.Column(6).Width;
                int rowHeight = 15;
                for (var i = 6; i <= MacData.Rows.Count + 6; i++)
                {
                    worksheetOA.Row(i).Height = rowHeight;
                }
                int PositionLooper = Row;
                Row = 6;
                int col = 10;
                int GraphsStartsAtRow = (rowHeight * MacData.Rows.Count) + 325;
                int PiesCount = 0;
                //PIE ForEach Machine :: Manual v/s DDL WO Count
                for (int n = 0; n < MacData.Rows.Count; n++)
                {
                    string machINV = Convert.ToString(MacData.Rows[n][1]);
                    //Facts: Starting Row, MacINV,WOCount,DDLCount Columns Names, No of Rows.
                    ValCells = null;
                    //PIE Chart for OverAll Manual v/s DDL WO Count
                    string ValueWO = Convert.ToString(worksheetOA.Cells["I" + (Row)].Value);
                    string ValueDDL = Convert.ToString(worksheetOA.Cells["J" + (Row)].Value);

                    if (ValueWO.Trim() != "0" && ValueDDL.Trim() != "0")
                    {
                        ValCells = "I" + (Row) + ",J" + (Row);
                        NameCells = "I5,J5";
                    }
                    else if (ValueDDL.Trim() != "0")
                    {
                        ValCells = "I" + (Row) + ",J" + (Row);
                        NameCells = "A1,J5";
                    }
                    else if (ValueWO.Trim() != "0")
                    {
                        ValCells = "I" + (Row) + ",J" + (Row);
                        NameCells = "I5,A1";
                    }

                    //NameCells = "I5,J5";


                    if (ValCells != null)
                    {
                        erValues = worksheetOA.Cells[ValCells];
                        erNames = worksheetOA.Cells[NameCells];
                        chartWOCount = (ExcelPieChart)worksheetOA.Drawings.AddChart("PieChartWOCount" + machINV, eChartType.Pie);
                        ExcelChart chartWOCount1 = (ExcelChart)chartWOCount;
                        chartWOCount.SetSize(300, 300);
                        //chartWOCount.SetPosition((PositionLooper + 8), 0, col, 0);
                        chartWOCount.SetPosition(GraphsStartsAtRow, col);
                        chartWOCount.Title.Text = machINV + "'s WO Count";
                        chartWOCount.Style = eChartStyle.Style18;
                        chartWOCount.Series.Add(erValues, erNames);
                        chartWOCount.DataLabel.ShowCategory = true;
                        chartWOCount.DataLabel.ShowPercent = true;
                        chartWOCount.Legend.Remove();
                        col += 300 + 10;
                        PiesCount++;

                        //trying to Put next set of Pie's below the current row.
                        //Logic: Add Another 310(PieSize + 10) Pixels to GraphsStartsAtRow && also the col Value.
                        //Question is : When.
                        if (PiesCount % 4 == 0)
                        {
                            GraphsStartsAtRow = GraphsStartsAtRow + 310;
                            col = 10;
                        }

                    }


                    Row++;
                }
            }


            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "ManualWOReport" + lowestLevelName + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "ManualWOReport" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                ////Response.Clear();
                ////Response.ClearContent();
                ////Response.ClearHeaders();
                ////Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                ////Response.AddHeader("Content-Length", file1.Length.ToString());
                ////Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                ////Response.WriteFile(file1.FullName);
                ////Response.Flush();
                ////Response.Close();
            }
            return path1;
        }

        //Utilization Report Method
        public async Task<string> UtilizationReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null, string TabularType = "Day")
        {
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\Utilization_Report.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateSummarized = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            string Header = null;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            var plantName = (from plant in db.tblplants
                                             where plant.PlantID == plantId
                                             select new { plantname = plant.PlantName }).SingleOrDefault();

                            Header = plantName.plantname;
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();

                        var shopName = (from shop in db.tblshops
                                        where shop.ShopID == shopId
                                        select new { shopname = shop.ShopName }).SingleOrDefault();
                        Header = shopName.shopname;

                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    var cellName = (from cell in db.tblcells
                                    where cell.CellID == cellId
                                    select new { wcname = cell.CellName }).SingleOrDefault();
                    Header = cellName.wcname;

                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                MacCount = 1;
                var WCName = (from wc in db.tblmachinedetails
                              where wc.MachineID == wcId
                              select new { wcname = wc.MachineDispName }).SingleOrDefault();
                Header = WCName.wcname;
            }

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    ////TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetOA = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetOA = p.Workbook.Worksheets.Add("Summarized", TemplateSummarized);
            }
            catch { }

            if (worksheet == null)
            {
                try
                {
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            } 
            if (worksheetOA == null)
            {
                try{
                worksheetOA = p.Workbook.Worksheets.Add("Summarized", TemplateSummarized);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //worksheet.Cells["C5"].Value = frmDate.ToString("dd-MM-yyyy");
            //worksheet.Cells["E5"].Value = toDate.ToString("dd-MM-yyyy");

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");

           

            DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime = frmDate.ToString("yyyy-MM-dd");
            DataTable machin = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = null;
                if (lowestLevel == "Plant")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where PlantID = " + PlantID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ); ";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where PlantID = " + PlantID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ); ";
                }
                else if (lowestLevel == "Shop")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where ShopID = " + ShopID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ); ";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where ShopID = " + ShopID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ); ";
                }
                else if (lowestLevel == "Cell")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where CellID = " + CellID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ); ";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where CellID = " + CellID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ); ";
                }
                else if (lowestLevel == "WorkCentre")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ); ";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails where MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ); ";
                }

                //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(machin);
                mc.close();
            }

            //DataTable for Consolidated Data 
            DataTable DTUtil = new DataTable();
            DTUtil.Columns.Add("Plant", typeof(string));
            DTUtil.Columns.Add("Shop", typeof(string));
            DTUtil.Columns.Add("Cell", typeof(string));
            DTUtil.Columns.Add("WC", typeof(string));
            DTUtil.Columns.Add("ValueAdding", typeof(double));
            DTUtil.Columns.Add("Setting", typeof(double));
            DTUtil.Columns.Add("Idle", typeof(double));
            DTUtil.Columns.Add("MinorLoss", typeof(double));
            DTUtil.Columns.Add("Breakdown", typeof(double));
            DTUtil.Columns.Add("NoPlan", typeof(double));
            DTUtil.Columns.Add("TotalDur", typeof(double));
            DTUtil.Columns.Add("Availability", typeof(double));

            //DataTable for Graph Data  (Daywise)
            DataTable DTUtilDayWise = new DataTable();
            // DTUtilDayWise.Columns.Add("Plant", typeof(string));
            //DTUtilDayWise.Columns.Add("Shop", typeof(string));
            //DTUtilDayWise.Columns.Add("Cell", typeof(string));
            //DTUtilDayWise.Columns.Add("WC", typeof(string));
            DTUtilDayWise.Columns.Add("Date", typeof(string));
            DTUtilDayWise.Columns.Add("ValueAdding", typeof(double));
            DTUtilDayWise.Columns.Add("Setting", typeof(double));
            DTUtilDayWise.Columns.Add("Idle", typeof(double));
            DTUtilDayWise.Columns.Add("MinorLoss", typeof(double));
            DTUtilDayWise.Columns.Add("Breakdown", typeof(double));
            DTUtilDayWise.Columns.Add("NoPlan", typeof(double));
            DTUtilDayWise.Columns.Add("TotalDur", typeof(double));
            DTUtilDayWise.Columns.Add("Utilization", typeof(double));


            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 7 + machin.Rows.Count + 2;
            int Sno = 1;

            //for (int i = 0; i < TotalDay + 1; i++)
            int l = 0;
            do
            {
                DateTime begining = UsedDateForExcel, end = UsedDateForExcel;
                var testDate = UsedDateForExcel;

                double DaysInCurrentPeriod = 0;
                if (TabularType == "Day")
                {
                    DaysInCurrentPeriod = 1;
                    begining = end = UsedDateForExcel;
                }
                else if (TabularType == "Week")
                {
                    GetWeek(testDate, new CultureInfo("fr-FR"), out begining, out end); //en-US(Sunday - Monday) //fr-FR(Monday - Sunday)
                    if (end.Subtract(toDate).TotalSeconds > 0)
                    {
                        end = toDate;
                    }
                    if (begining.Subtract(UsedDateForExcel).TotalSeconds < 0)
                    {
                        begining = UsedDateForExcel;
                    }
                    DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                }
                else if (TabularType == "Month")
                {
                    DateTime itempDate = UsedDateForExcel.AddMonths(1);
                    DateTime Temp2 = new DateTime(itempDate.Year, itempDate.Month, 01, 00, 00, 01);
                    end = Temp2.AddDays(-1);
                    if (end.Subtract(toDate).TotalSeconds > 0)
                    {
                        end = toDate;
                    }
                    DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                }
                else if (TabularType == "Year")
                {
                    DateTime Temp2 = new DateTime(UsedDateForExcel.AddYears(1).Year, 01, 01, 00, 00, 01);
                    end = Temp2.AddDays(-1);
                    if (end.Subtract(toDate).TotalSeconds > 0)
                    {
                        end = toDate;
                    }
                    DaysInCurrentPeriod = end.Subtract(UsedDateForExcel).TotalDays + 1;
                }

                int thisDatesDataStartsAt = Row;
                double IndividualDateValueAdding = 0, IndividualDateSetting = 0, IndividualDateIdle = 0, IndividualDateMinorLoss = 0, IndividualDateBreakdown = 0, IndividualDateNoPlan = 0, IndividualDateTotal = 0;
                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");
                string correctedDateS = UsedDateForExcel.ToString("yyyy-MM-dd");
                string correctedDateE = UsedDateForExcel.ToString("yyyy-MM-dd");
                if (TabularType != "Day")
                {
                    if (end.Subtract(toDate).TotalSeconds < 0)
                    {
                        correctedDateE = end.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        correctedDateE = toDate.ToString("yyyy-MM-dd");
                    }
                }

                for (int n = 0; n < machin.Rows.Count; n++)
                {
                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    double AvaillabilityFactor, EfficiencyFactor, QualityFactor;
                    double green, red, yellow, blue = 0, MinorLoss = 0, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
                    double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
                    double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;

                    AvailableTime = 24 * 60; //24Hours to Minutes
                    double Diff = 0;

                    // 2017-02-21
                    //MinorLosses = GetMinorLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
                    //blue = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");

                    ////Availability
                    //SettingTime = GetSettingTime(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                    //ROALossess = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROA");
                    //DownTimeBreakdown = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
                    ////DownTimeBreakdown = GetDownTimeBreakdown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                    ////OperatingTime = AvailableTime - (ROALossess + DownTimeBreakdown + blue + MinorLosses);

                    //green = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
                    //OperatingTime = green;
                    // 2017-02-21

                    //New Logic . Take values from 

                    //var OEEData = db.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                    //if (OEEData != null)
                    //{
                    //    MinorLosses = Convert.ToDouble(OEEData.MinorLosses);
                    //    blue = Convert.ToDouble(OEEData.Blue);
                    //    SettingTime = Convert.ToDouble(OEEData.SettingTime);
                    //    ROALossess = Convert.ToDouble(OEEData.ROALossess);
                    //    DownTimeBreakdown = Convert.ToDouble(OEEData.DownTimeBreakdown);
                    //    //SummationOfSCTvsPP = Convert.ToDouble(OEEData.SummationOfSCTvsPP);
                    //    green = Convert.ToDouble(OEEData.Green);
                    //    OperatingTime = green;
                    //    ScrapQtyTime = Convert.ToDouble(OEEData.ScrapQtyTime);
                    //    ReWOTime = Convert.ToDouble(OEEData.ReWOTime);
                    //}
                    //else
                    //{
                    #region Calling .exe Not Implemented
                    //    //try
                    //    //{
                    //    //Server.MapPath(@"\Content\DATA"); 

                    //    String cPath = db.tblapp_paths.Where(m => m.IsDeleted == 0 && m.AppName == "CalOEEDaily").Select(m => m.AppPath).FirstOrDefault();
                    //    //string filename = Path.Combine(@"\Content1\Debug\", "CalOEEDaily.exe");
                    //    //var proc = System.Diagnostics.Process.Start(filename, DateTimeValue.ToString());
                    //    //proc.WaitForExit();
                    //    //proc.Kill();
                    //    try
                    //    {
                    //        string filename = Path.Combine(cPath, "CalOEEDaily.exe");
                    //        var proc = System.Diagnostics.Process.Start(Server.MapPath(@filename), DateTimeValue.ToString());
                    //        proc.WaitForExit();
                    //    }
                    //    catch (Exception e)
                    //    {
                    //        FileInfo fi = new FileInfo(Path.Combine(cPath, "CalOEEDaily.exe"));

                    //    }
                    //    // proc.Kill();

                    //    ////Try this 2017-03-01
                    //    //System.Security.Principal.WindowsImpersonationContext impersonationContext;
                    //    //impersonationContext =
                    //    //    ((System.Security.Principal.WindowsIdentity)User.Identity).Impersonate();

                    //    ////Insert your code that runs under the security context of the authenticating user here.
                    //    //String cPath = db.tblapp_paths.Where(m => m.IsDeleted == 0 && m.AppName == "CalOEEDaily").Select(m => m.AppPath).FirstOrDefault();
                    //    //string filename = Path.Combine(cPath, "CalOEEDaily.exe");
                    //    //var proc = System.Diagnostics.Process.Start(filename, DateTimeValue.ToString());
                    //    //proc.WaitForExit();
                    //    //impersonationContext.Undo();
                    //    //}
                    //    //catch (Exception e)
                    //    //{

                    //    //}

                    //    var OEEDataInner = db.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                    //    if (OEEDataInner != null)
                    //    {
                    //        MinorLosses = Convert.ToDouble(OEEDataInner.MinorLosses);
                    //        blue = Convert.ToDouble(OEEDataInner.Blue);
                    //        SettingTime = Convert.ToDouble(OEEDataInner.SettingTime);
                    //        ROALossess = Convert.ToDouble(OEEDataInner.ROALossess);
                    //        DownTimeBreakdown = Convert.ToDouble(OEEDataInner.DownTimeBreakdown);
                    //        //SummationOfSCTvsPP = Convert.ToDouble(OEEData.SummationOfSCTvsPP);
                    //        green = Convert.ToDouble(OEEDataInner.Green);
                    //        OperatingTime = green;
                    //        ScrapQtyTime = Convert.ToDouble(OEEDataInner.ScrapQtyTime);
                    //        ReWOTime = Convert.ToDouble(OEEDataInner.ReWOTime);
                    //    }
                    //    else
                    //    {
                    //        continue;
                    //    }
                    #endregion
                    //}

                    DateTime DateTimeValue = Convert.ToDateTime(UsedDateForExcel.ToString("yyyy-MM-dd") + " " + "00:00:00");
                    DataTable OEEData = new DataTable();
                    using (MsqlConnection mcOEE = new MsqlConnection())
                    {
                        mcOEE.open();
                        string query = null;
                        query = @"select WCID,sum(Green), sum(SummationOfSCTvsPP),sum(SettingTime),sum(ROALossess),SUM(MinorLosses), sum(Blue), sum(DownTimeBreakdown), "
                                + " sum(ScrapQtyTime),sum(ReWOTime) from i_facility_tal.dbo.tbloeedashboardvariables where StartDate >= '" + correctedDateS + "' and StartDate <= '" + correctedDateE + "' and WCID = '" + MachineID + "' group by WCID ;";
                        //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                        SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcOEE.msqlConnection);
                        daLossCodesData.Fill(OEEData);
                        mcOEE.close();
                    }
                    if (OEEData.Rows.Count > 0)
                    {
                        green = Convert.ToDouble(OEEData.Rows[0][1]);
                        SummationOfSCTvsPP = Convert.ToDouble(OEEData.Rows[0][2]);
                        SettingTime = Convert.ToDouble(OEEData.Rows[0][3]);
                        MinorLosses = Convert.ToDouble(OEEData.Rows[0][5]);
                        ROALossess = Convert.ToDouble(OEEData.Rows[0][4]);
                        blue = Convert.ToDouble(OEEData.Rows[0][6]);
                        DownTimeBreakdown = Convert.ToDouble(OEEData.Rows[0][7]);
                        ScrapQtyTime = Convert.ToDouble(OEEData.Rows[0][8]);
                        ReWOTime = Convert.ToDouble(OEEData.Rows[0][9]);
                        OperatingTime = green;
                    }

                    double TotalMinutes = OperatingTime + (SettingTime) + (ROALossess - SettingTime) + MinorLosses + DownTimeBreakdown + blue;
                    //double Diff = 1440 - TotalMinutes;
                    Double ActualTotalMinutes = (8 * 3 * 60) * DaysInCurrentPeriod;
                    Diff = ActualTotalMinutes - TotalMinutes;
                    if (Diff > 0)
                    {
                        blue += Diff;
                    }
                    else
                    {
                        ROALossess += Diff;
                    }

                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                    worksheet.Cells["D" + Row].Value = HierarchyData[0];
                    worksheet.Cells["E" + Row].Value = HierarchyData[1];
                    worksheet.Cells["F" + Row].Value = HierarchyData[2];
                    worksheet.Cells["G" + Row].Value = HierarchyData[3];

                    double ValueAddingTime = Math.Round(OperatingTime, 2);
                    double setTime = Math.Round(SettingTime, 2);
                    double idleTime = Math.Round((ROALossess - SettingTime), 2);
                    double minorLossTime = Math.Round(MinorLosses, 2);
                    double BreakdownTime = Math.Round(DownTimeBreakdown, 2);
                    double blueTime = Math.Round(blue, 2);

                    //For Individual Date Cummulative
                    IndividualDateValueAdding += ValueAddingTime;
                    IndividualDateSetting += setTime;
                    IndividualDateIdle += idleTime;
                    IndividualDateMinorLoss += minorLossTime;
                    IndividualDateBreakdown += BreakdownTime;
                    IndividualDateNoPlan += blueTime;

                    worksheet.Cells["H" + Row].Value = Math.Round((OperatingTime / 60), 1);
                    worksheet.Cells["I" + Row].Value = Math.Round((SettingTime / 60), 1);
                    worksheet.Cells["J" + Row].Value = Math.Round(((ROALossess - SettingTime) / 60), 1);
                    worksheet.Cells["K" + Row].Value = Math.Round((MinorLosses / 60), 1);
                    worksheet.Cells["L" + Row].Value = Math.Round((DownTimeBreakdown / 60), 1);
                    worksheet.Cells["M" + Row].Value = Math.Round((blue / 60), 1);

                    double IntermediateTotal = Math.Round(OperatingTime, 2) + Math.Round(SettingTime, 2) + Math.Round((ROALossess - SettingTime), 2) + Math.Round(MinorLosses, 2) + Math.Round(DownTimeBreakdown, 2) + Math.Round(blue, 2);
                    IndividualDateTotal += IntermediateTotal;
                    double TotalHours = (double)worksheet.Calculate("=SUM(H" + Row + ":M" + Row + ")");
                    worksheet.Cells["N" + Row].Formula = Convert.ToString(Math.Round(TotalHours, 0));

                    string MacName = Convert.ToString(HierarchyData[3]);
                    //MacName.Replace("-","\\-");
                    DataRow dr = DTUtil.Select("WC = '" + @MacName + "'").FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                    if (dr != null)
                    {
                        double ValueAddPrev = Convert.ToDouble(dr["ValueAdding"]);
                        dr["ValueAdding"] = (ValueAddPrev + ValueAddingTime);

                        double setPrev = Convert.ToDouble(dr["Setting"]);
                        dr["Setting"] = (setPrev + setTime);

                        double idlePrev = Convert.ToDouble(dr["Idle"]);
                        dr["Idle"] = (idlePrev + idleTime);

                        double MinorLPrev = Convert.ToDouble(dr["MinorLoss"]);
                        dr["MinorLoss"] = (MinorLPrev + minorLossTime);

                        double BreakPrev = Convert.ToDouble(dr["Breakdown"]);
                        dr["Breakdown"] = (BreakPrev + BreakdownTime);

                        double bluePrev = Convert.ToDouble(dr["NoPlan"]);
                        dr["NoPlan"] = (bluePrev + blueTime);

                        double TotalDurPrev = Convert.ToDouble(dr["TotalDur"]);
                        dr["TotalDur"] = (TotalDurPrev + IntermediateTotal);

                    }
                    else
                    {
                        DTUtil.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3], ValueAddingTime, setTime, idleTime, minorLossTime, BreakdownTime, blueTime, IntermediateTotal);
                    }

                    //Availability // Whole day duration is 24*60 = 1440 minutes
                    double val = OperatingTime / AvailableTime;
                    if (val > 0)
                    {
                        worksheet.Cells["O" + Row].Value = Math.Round(val * 100, 0);
                    }
                    else
                    {
                        worksheet.Cells["O" + Row].Value = 0;
                    }

                    Row++;
                    Sno++;
                    worksheet.Row(Row).CustomHeight = false;
                }

                if (lowestLevel != "WorkCentre")
                {
                    worksheet.Cells["H" + Row].Value = Math.Round(IndividualDateValueAdding / 60, 1);
                    worksheet.Cells["I" + Row].Value = Math.Round(IndividualDateSetting / 60, 1);
                    worksheet.Cells["J" + Row].Value = Math.Round(IndividualDateIdle / 60, 1);
                    worksheet.Cells["K" + Row].Value = Math.Round(IndividualDateMinorLoss / 60, 1);
                    worksheet.Cells["L" + Row].Value = Math.Round(IndividualDateBreakdown / 60, 1);
                    worksheet.Cells["M" + Row].Value = Math.Round(IndividualDateNoPlan / 60, 1);
                    worksheet.Cells["N" + Row].Value = Math.Round(IndividualDateTotal / 60, 1);
                    var blah = Math.Round((IndividualDateValueAdding / (24 * 60 * machin.Rows.Count)) * 100, 1);
                    worksheet.Cells["O" + Row].Value = Math.Round((IndividualDateValueAdding / (24 * 60 * machin.Rows.Count)) * 100, 0);
                }
                //Add data into datatable for graph
                DTUtilDayWise.Rows.Add(UsedDateForExcel.ToString("yyyy-MM-dd"),
                    Math.Round(IndividualDateValueAdding / 60, 1),
                    Math.Round(IndividualDateSetting / 60, 1),
                    Math.Round(IndividualDateIdle / 60, 1),
                    Math.Round(IndividualDateMinorLoss / 60, 1),
                    Math.Round(IndividualDateBreakdown / 60, 1),
                    Math.Round(IndividualDateNoPlan / 60, 1),
                     Math.Round(IndividualDateTotal / 60, 1),
                    Math.Round((IndividualDateValueAdding / (24 * 60 * machin.Rows.Count)), 1));

                #region Individual Date Cummulative
                if (Row > thisDatesDataStartsAt)
                {
                    //thin border for all cells
                    var modelTable = worksheet.Cells["B" + thisDatesDataStartsAt + ":O" + Row];
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Insert OverAll NoLogin Time
                    worksheet.Cells["B" + Row + ":G" + Row].Merge = true;
                    worksheet.Cells["B" + Row].Value = "Duration";
                    worksheet.Cells["B" + Row + ":O" + Row].Style.Font.Bold = true;
                    worksheet.Cells["B" + Row + ":O" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    worksheet.Cells["H" + Row].Formula = "=SUM(H" + thisDatesDataStartsAt + ":H" + (Row - 1) + ")";
                    worksheet.Cells["I" + Row].Formula = "=SUM(I" + thisDatesDataStartsAt + ":I" + (Row - 1) + ")";
                    worksheet.Cells["J" + Row].Formula = "=SUM(J" + thisDatesDataStartsAt + ":J" + (Row - 1) + ")";
                    worksheet.Cells["K" + Row].Formula = "=SUM(K" + thisDatesDataStartsAt + ":K" + (Row - 1) + ")";
                    worksheet.Cells["L" + Row].Formula = "=SUM(L" + thisDatesDataStartsAt + ":L" + (Row - 1) + ")";
                    worksheet.Cells["M" + Row].Formula = "=SUM(M" + thisDatesDataStartsAt + ":M" + (Row - 1) + ")";
                    worksheet.Cells["N" + Row].Formula = "=SUM(N" + thisDatesDataStartsAt + ":N" + (Row - 1) + ")";

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + thisDatesDataStartsAt + ":O" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    Row++;

                }

                #endregion

                double offsetE = 1;
                if (TabularType == "Day")
                {
                    offsetE = 1;
                }
                else if (TabularType == "Week")
                {
                    offsetE = end.Subtract(begining).TotalDays + 1;
                }
                else if (TabularType == "Month")
                {
                    offsetE = end.Subtract(begining).TotalDays + 1;
                }
                else if (TabularType == "Year")
                {
                    offsetE = end.Subtract(begining).TotalDays + 1;
                }
                UsedDateForExcel = UsedDateForExcel.AddDays(offsetE);
                Row++;
                //UsedDateForExcel = UsedDateForExcel.AddDays(+1);

            } while (UsedDateForExcel <= toDate);

            Sno = 1;
            Row = 6 + 1;
            //push Cummulative Data
            for (int i = 0; i < DTUtil.Rows.Count; i++)
            {
                worksheet.Cells["B" + Row].Value = Sno;
                worksheet.Cells["C" + Row].Value = frmDate.ToString("yyyy-MM-dd") + " To " + toDate.ToString("yyyy-MM-dd");
                worksheet.Cells["D" + Row].Value = DTUtil.Rows[i][0];
                worksheet.Cells["E" + Row].Value = DTUtil.Rows[i][1];
                worksheet.Cells["F" + Row].Value = DTUtil.Rows[i][2];
                worksheet.Cells["G" + Row].Value = DTUtil.Rows[i][3];

                double valueAddingSummation1 = Convert.ToDouble(DTUtil.Rows[i][4]);
                double setTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][5]);
                double idleTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][6]);
                double minorLossTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][7]);
                double BreakdownTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][8]);
                double blueTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][9]);
                double TotalTimeSummation1 = Convert.ToDouble(DTUtil.Rows[i][10]);

                worksheet.Cells["H" + Row].Value = Math.Round(valueAddingSummation1 / (60), 1);
                worksheet.Cells["I" + Row].Value = Math.Round(setTimeSummation1 / (60), 1);
                worksheet.Cells["J" + Row].Value = Math.Round(idleTimeSummation1 / (60), 1);
                worksheet.Cells["K" + Row].Value = Math.Round(minorLossTimeSummation1 / (60), 1);
                worksheet.Cells["L" + Row].Value = Math.Round(BreakdownTimeSummation1 / (60), 1);
                worksheet.Cells["M" + Row].Value = Math.Round(blueTimeSummation1 / (60), 1);
                worksheet.Cells["N" + Row].Value = Math.Round(TotalTimeSummation1 / (60), 1);

                worksheet.Cells["O" + Row].Value = Math.Round((Convert.ToDouble(DTUtil.Rows[i][4]) / (24 * 60 * (TotalDay + 1))) * 100, 0);

                Sno++;
                Row++;
            }

            //Summarized (Selected)TopMostLevel
            //Row = 7;
            double valueAddingSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("ValueAdding"));
            double setTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("Setting"));
            double idleTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("Idle"));
            double minorLossTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("MinorLoss"));
            double BreakdownTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("Breakdown"));
            double blueTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("NoPlan"));
            double TotalTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("TotalDur"));

            worksheet.Cells["H" + Row].Value = Math.Round(valueAddingSummation / (60), 1);
            worksheet.Cells["I" + Row].Value = Math.Round(setTimeSummation / (60), 1);
            worksheet.Cells["J" + Row].Value = Math.Round(idleTimeSummation / (60), 1);
            worksheet.Cells["K" + Row].Value = Math.Round(minorLossTimeSummation / (60), 1);
            worksheet.Cells["L" + Row].Value = Math.Round(BreakdownTimeSummation / (60), 1);
            worksheet.Cells["M" + Row].Value = Math.Round(blueTimeSummation / (60), 1);
            worksheet.Cells["N" + Row].Value = Math.Round(TotalTimeSummation / (60), 1);
            worksheet.Cells["O" + Row].Value = Math.Round((valueAddingSummation / ((TotalDay + 1) * 24 * 60 * machin.Rows.Count)) * 100, 0);

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            var modelTableUtil = worksheet.Cells["B6:O" + Row];
            // Assign borders
            modelTableUtil.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTableUtil.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTableUtil.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTableUtil.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            worksheet.Cells["B" + Row + ":G" + Row].Merge = true;
            worksheet.Cells["B" + Row].Value = " Total Duration";
            worksheet.Cells["B" + Row + ":O" + Row].Style.Font.Bold = true;
            worksheet.Cells["B" + Row + ":O" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            //Excel:: Border Around Cells.
            worksheet.Cells["B6:O" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            //For Header
            worksheet.Cells["F2:L2"].Merge = true;
            worksheet.Cells["F2:I2"].Style.Font.Bold = true;
            worksheet.Cells["F2:I2"].Style.Font.Size = 16;
            worksheet.Cells["F2"].Value = Header.ToUpper() + " Utilization Report";
            worksheet.Cells["F2:L2"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
            worksheetOA.Cells["F2:L2"].Merge = true;
            worksheetOA.Cells["F2"].Style.Font.Bold = true;
            worksheetOA.Cells["F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheetOA.Cells["F2"].Style.Font.Size = 16;
            worksheetOA.Cells["F2"].Value = Header.ToUpper() + " Utilization Report";
            worksheetOA.Cells["F2:L2"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            worksheet.View.ShowGridLines = false;
            worksheetOA.View.ShowGridLines = false;
            worksheet.Cells["B" + 6 + ":O" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            #region The Graph
            int row = 5;
            //Below u r just adding 1 bar (stacked bar with all Data) per day into same graph
            if (DTUtilDayWise.Rows.Count > 0)
            {
                worksheetOA.Cells["C" + row].Value = "Value Adding";
                worksheetOA.Cells["D" + row].Value = "Setting";
                worksheetOA.Cells["E" + row].Value = "Idle";
                worksheetOA.Cells["F" + row].Value = "Minor Loss";
                worksheetOA.Cells["G" + row].Value = "Breakdown";
                worksheetOA.Cells["H" + row].Value = "No Plan";

                for (int i = 0; i < DTUtilDayWise.Rows.Count; i++)
                {
                    row++;
                    double OneDaysTotal = (double)DTUtilDayWise.Rows[i][7];
                    worksheetOA.Cells["B" + row].Value = DTUtilDayWise.Rows[i][0];
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][1]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["C" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][1]) / OneDaysTotal) * 100, 1);
                    }
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][2]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["D" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][2]) / OneDaysTotal) * 100, 1);
                    }
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][3]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["E" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][3]) / OneDaysTotal) * 100, 1);
                    }
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][4]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["F" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][4]) / OneDaysTotal) * 100, 1);
                    }
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][5]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["G" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][5]) / OneDaysTotal) * 100, 1);
                    }
                    if (Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][6]) / OneDaysTotal) * 100, 1) > 0)
                    {
                        worksheetOA.Cells["H" + row].Value = Math.Round((Convert.ToDouble(DTUtilDayWise.Rows[i][6]) / OneDaysTotal) * 100, 1);
                    }



                }
                //worksheetOA.Cells["A1"].Value = "Ratio of Losses";
                ExcelRange erDateNAMES = worksheetOA.Cells["B6:B" + row];

                ExcelRange erValueAddingValues = worksheetOA.Cells["C6:C" + row];
                double ValueAddingSum = (double)worksheetOA.Calculate("=SUM(" + erValueAddingValues + ")");

                ExcelRange erSettingValues = worksheetOA.Cells["D6:D" + row];
                ExcelRange erIdleValues = worksheetOA.Cells["E6:E" + row];
                ExcelRange erMinorLossValues = worksheetOA.Cells["F6:F" + row];
                ExcelRange erBreakdownValues = worksheetOA.Cells["G6:G" + row];
                ExcelRange erNoPlanValues = worksheetOA.Cells["H6:H" + row];

                var chartIDAndUnID = (ExcelBarChart)worksheetOA.Drawings.AddChart("Utilization", eChartType.ColumnStacked);
                if (DTUtilDayWise.Rows.Count < 10)
                {
                    chartIDAndUnID.SetSize(500, 350);
                }
                else
                {
                    chartIDAndUnID.SetSize((27 * DTUtilDayWise.Rows.Count), 350);
                }
                chartIDAndUnID.SetPosition(75, 20);

                chartIDAndUnID.Title.Text = "Utilization";
                chartIDAndUnID.Style = eChartStyle.Style18;
                chartIDAndUnID.Legend.Position = eLegendPosition.Bottom;
                //chartIDAndUnID.Legend.Remove();
                chartIDAndUnID.YAxis.MaxValue = 100;
                chartIDAndUnID.YAxis.MinValue = 0;
                chartIDAndUnID.Locked = false;
                chartIDAndUnID.PlotArea.Border.Width = 0;
                chartIDAndUnID.YAxis.MinorTickMark = eAxisTickMark.None;
                chartIDAndUnID.DataLabel.ShowValue = true;
                chartIDAndUnID.DisplayBlanksAs = eDisplayBlanksAs.Gap;
                var ValueAdding = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erValueAddingValues, erDateNAMES));
                ValueAdding.Header = "ValueAdding";
                var Setting = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erSettingValues, erDateNAMES));
                Setting.Header = "Setting";
                var Idle = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erIdleValues, erDateNAMES));
                Idle.Header = "Idle";
                var MinorLoss = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erMinorLossValues, erDateNAMES));
                MinorLoss.Header = "MinorLoss";
                var Breakdown = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erBreakdownValues, erDateNAMES));
                Breakdown.Header = "Breakdown";
                var NoPlan = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erNoPlanValues, erDateNAMES));
                NoPlan.Header = "NoPlan";

                //Append the Avg of ValueAdding to Chart.
                //populate excel with AvgValue of ValueAdding
                string Col1 = "K", Col2 = "L";
                double ValueAddingAvg = ValueAddingSum / DTUtilDayWise.Rows.Count;
                for (int i = 0; i < DTUtilDayWise.Rows.Count; i++)
                {
                    worksheetOA.Cells[Col1 + (50 + i)].Value = "Avg";
                    worksheetOA.Cells[Col2 + (50 + i)].Value = ValueAddingAvg;
                }
                var chartTypeL1 = chartIDAndUnID.PlotArea.ChartTypes.Add(eChartType.Line);
                var serieL2 = chartTypeL1.Series.Add(worksheetOA.Cells[Col2 + "50:" + Col2 + (50 + DTUtilDayWise.Rows.Count - 1)], worksheetOA.Cells[Col1 + "50:" + Col1 + (50 + DTUtilDayWise.Rows.Count - 1)]);



                //Get reference to the worksheet xml for proper namespace
                var chartXml = chartIDAndUnID.ChartXml;
                var nsuri = chartXml.DocumentElement.NamespaceURI;
                var nsm = new XmlNamespaceManager(chartXml.NameTable);
                nsm.AddNamespace("c", nsuri);

                //XY Scatter plots have 2 value axis and no category
                var valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
                if (valAxisNodes != null && valAxisNodes.Count > 0)
                    foreach (XmlNode valAxisNode in valAxisNodes)
                    {
                        var major = valAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                        if (major != null)
                            valAxisNode.RemoveChild(major);

                        var minor = valAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                        if (minor != null)
                            valAxisNode.RemoveChild(minor);
                    }

                //Other charts can have a category axis
                var catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
                if (catAxisNodes != null && catAxisNodes.Count > 0)
                    foreach (XmlNode catAxisNode in catAxisNodes)
                    {
                        var major = catAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                        if (major != null)
                            catAxisNode.RemoveChild(major);

                        var minor = catAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                        if (minor != null)
                            catAxisNode.RemoveChild(minor);
                    }


            }



            #endregion

            //worksheetOA.Cells["A3:R100"].Style.Font.Color.SetColor(Color.White);
            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "Utilization_Report" + Header.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "Utilization_Report " + Header + " " + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                ////Response.Clear();
                ////Response.ClearContent();
                ////Response.ClearHeaders();
                ////Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                ////Response.AddHeader("Content-Length", file1.Length.ToString());
                ////Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                ////Response.WriteFile(file1.FullName);
                ////Response.Flush();
                ////Response.Close();
            }
            return path1;
        }

        //NoLogin Report Method
        public async Task<string> NoLoginReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\NoLoginReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateSummarized = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            string lowestLevel = null;
            string lowestLevelName = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "NoLoginReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "NoLoginReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetOA = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            } if (worksheetOA == null)
            {
                try{
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //Step1: Get the Summarized according to Name. for each WC from StartDate - EndDate
            //( Leave that many rows blank and Fill WorkCenter wise 1st.
            

            //DataTable for Consolidated Duration 
            DataTable DTNoLogin = new DataTable();
            DTNoLogin.Columns.Add("MachineID", typeof(int));
            DTNoLogin.Columns.Add("Duration", typeof(int));


            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5;
            int Sno = 1;
            for (int i = 0; i < TotalDay + 1; i++)
            {
                DateTime endDateTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = UsedDateForExcel.ToString("yyyy-MM-dd");

                DataTable HMIData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = null;
                    if (lowestLevel == "Plant")
                    {
                        //query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "Shop")
                    {
                        //query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "Cell")
                    {
                        //query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "WorkCentre")
                    {
                        //query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, Min(PEStartTime) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    //SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(HMIData);
                    mc.close();
                }
                double individualDateDuration = 0;
                int thisDatesDataStartsAt = Row;

                for (int n = 0; n < HMIData.Rows.Count; n++)
                {
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        thisDatesDataStartsAt = Row;
                    }
                    int MachineID = Convert.ToInt32(HMIData.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = HierarchyData[3];

                    //string Duration = null;
                    double DurInSeconds = 0;
                    string date = UsedDateForExcel.ToString("yyyy-MM-dd");
                    DateTime DayStartDateTime = Convert.ToDateTime(date) + new TimeSpan(06, 00, 00);
                    if (!string.IsNullOrEmpty(Convert.ToString(HMIData.Rows[n][1])))
                    {
                        DateTime LoginTime = Convert.ToDateTime(HMIData.Rows[n][1]);
                        DurInSeconds = Convert.ToInt32(LoginTime.Subtract(DayStartDateTime).TotalSeconds);
                    }
                    //Duration = TimeFromSeconds(DurInSeconds);

                    individualDateDuration += Convert.ToDouble(DurInSeconds);

                    worksheet.Cells["G" + Row].Value = DayStartDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells["H" + Row].Value = Convert.ToDateTime(HMIData.Rows[n][1]).ToString("yyyy-MM-dd HH:mm:ss");

                    worksheet.Cells["I" + Row].Value = Math.Round((DurInSeconds / 3600), 2);

                    //worksheet.Cells["I" + Row].Value = Duration;

                    DataRow dr = DTNoLogin.Select("MachineID = " + MachineID).FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                    if (dr != null)
                    {
                        Int64 DurationPrev = Convert.ToInt64(dr["Duration"]); //get lossduration and update it.
                        dr["Duration"] = (DurationPrev + DurInSeconds);
                    }
                    else
                    {
                        DTNoLogin.Rows.Add(MachineID, DurInSeconds);
                    }
                    //double Minutes = Math.Round(Convert.ToDateTime(HMIData.Rows[n][5]).Subtract(Convert.ToDateTime(HMIData.Rows[n][4])).TotalMinutes, 2);

                    Row++;
                }

                if (Row > thisDatesDataStartsAt)
                {
                    //thin border for all cells
                    var modelTable = worksheet.Cells["B" + thisDatesDataStartsAt + ":I" + Row];
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //thick border around cells
                    // worksheet.Cells["B"+ thisDatesDataStartsAt +":I" +Row].BorderAround();
                    //worksheet.Cells[8, i].Style.Border.BorderAround["B" + thisDatesDataStartsAt + ":I" + Row] = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); 

                    //Insert OverAll NoLogin Time
                    worksheet.Cells["B" + Row + ":H" + Row].Merge = true;
                    worksheet.Cells["B" + Row].Value = "Duration";
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Font.Bold = true;
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    double SumNoLogin = 0;
                    SumNoLogin = Convert.ToDouble(individualDateDuration / 3600);
                    double RoundOffSumNoLogin = Math.Round(SumNoLogin, 1);
                    worksheet.Cells["I" + Row].Value = RoundOffSumNoLogin;

                    // worksheet.Cells["I" + Row].Value = Convert.ToString(TimeFromSeconds(individualDateDuration));

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + thisDatesDataStartsAt + ":I" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                }
                Row++;
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }
            //to Align Center for sheet detailed
            worksheet.Cells["B" + 5 + ":I" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            #region To Push Summarized Data

            Row = 6;
            Sno = 1;

            //sorted Details
            DataView dv = DTNoLogin.DefaultView;
            dv.Sort = "Duration desc";
            DataTable SortedDTNoLogin = dv.ToTable();

            double SummarizedNoLoginDuration = 0;
            for (int n = 0; n < SortedDTNoLogin.Rows.Count; n++)
            {
                int MachineID = Convert.ToInt32(SortedDTNoLogin.Rows[n][0]);
                List<string> HierarchyData = GetHierarchyData(MachineID);

                worksheetOA.Cells["B" + Row].Value = Sno++;
                worksheetOA.Cells["C" + Row].Value = HierarchyData[0];
                worksheetOA.Cells["D" + Row].Value = HierarchyData[1];
                worksheetOA.Cells["E" + Row].Value = HierarchyData[2];
                worksheetOA.Cells["F" + Row].Value = HierarchyData[3];

                worksheetOA.Cells["G" + Row].Value = Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd");
                worksheetOA.Cells["H" + Row].Value = Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd");

                string Duration = "0"; double Min = 0; double Min1 = 0;
                DataRow dr = SortedDTNoLogin.Select("MachineID = " + MachineID).FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                if (dr != null)
                {
                    Duration = Convert.ToString(dr["Duration"]); //get lossduration and update it.
                }
                SummarizedNoLoginDuration += Convert.ToDouble(Duration);

                // before convertion passing into Mitutes
                Min = Convert.ToDouble(Duration);
                Min1 = Convert.ToDouble(Min / 3600); //convert into minutes

                Duration = TimeFromSeconds(Convert.ToInt64(Duration));
                // worksheetOA.Cells["I" + Row].Value = Duration;

                double RoundOffMin1 = Math.Round(Min1, 1);
                worksheetOA.Cells["Z" + Row].Value = RoundOffMin1;

                worksheetOA.Cells["I" + Row].Value = RoundOffMin1;

                //Making Text white
                //Color ColorHexWhite = System.Drawing.Color.White;
                //worksheetOA.Cells["Z6:Z" + Row + ""].Style.Font.Color.SetColor(ColorHexWhite);

                Row++;
            }

            //thin border for all cells
            var modelTable1 = worksheetOA.Cells["B5:I" + Row];
            // Assign borders
            modelTable1.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //Row = 6;
            //Row++;
            worksheetOA.Cells["B" + Row + ":H" + Row].Merge = true;
            worksheetOA.Cells["B" + Row].Value = "Total Duration";
            worksheetOA.Cells["B" + Row + ":I" + Row].Style.Font.Bold = true;
            worksheetOA.Cells["B" + Row + ":I" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            double SumNoLogin1 = 0;
            SumNoLogin1 = Convert.ToDouble(SummarizedNoLoginDuration / 3600);
            double RoundOffSumNoLogin1 = Math.Round(SumNoLogin1, 1);
            worksheetOA.Cells["I" + Row].Value = RoundOffSumNoLogin1;

            //worksheetOA.Cells["I" + Row].Value = TimeFromSeconds(Convert.ToInt64(SummarizedNoLoginDuration));


            //Excel:: Border Around Cells.
            worksheetOA.Cells["B6:I" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

            #endregion
            try
            {
                worksheetOA.Cells["B" + 4 + ":I" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheetOA.Cells["B" + 5 + ":I" + Row].AutoFitColumns();
                //worksheetOA.Cells["B" + 5 + ":I" + 5].AutoFitColumns();
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            }
            catch (Exception e)
            {
            }
            //worksheetOA.Cells["B" + 5 + ":I" + Row].AutoFitColumns();
            worksheet.View.ShowGridLines = false;
            worksheetOA.View.ShowGridLines = false;

            try
            {
                #region The Graph, Pareto Graph for No Login Time
                if (Row != 6)
                {
                    Row = Row - 1;

                    ExcelRange MachinesNames = worksheetOA.Cells["F6:F" + Row];
                    ExcelRange WODuration = worksheetOA.Cells["Z6:Z" + Row];

                    var WOChart = (ExcelBarChart)worksheetOA.Drawings.AddChart("Summarized" + System.DateTime.Now.ToString("dd-MM-yyyy"), eChartType.ColumnClustered);
                    ExcelChart WOChartExcel = (ExcelBarChart)WOChart;
                    #region To calculate the Auto width and Height for Graph position
                    Row = Row + 3;
                    int row = worksheetOA.Cells["A1:A" + Row].Rows;// cellAddress.Start.Row;
                    int column = worksheetOA.Cells["I1:I" + Row].Columns;// cellAddress.Start.Column;

                    int rowSpan = row;
                    int colSpan = column;

                    double height = 0, width = 0;
                    for (int h = 0; h < rowSpan; h++)
                    {
                        height += worksheetOA.Row(row + h).Height;
                    }
                    for (int w = 0; w < colSpan; w++)
                    {
                        width += worksheetOA.Column(column + w).Width;
                    }

                    double pointToPixel = 0.75;

                    height /= pointToPixel;
                    width /= 0.1423;
                    WOChart.SetPosition((int)height, (int)width);

                    #endregion End Auto width and Height graph position

                    WOChart.SetSize(600, 300);
                    WOChart.Title.Text = "No Login Time";
                    WOChart.Style = eChartStyle.Style15;
                    WOChart.Legend.Position = eLegendPosition.Right;
                    //WOChart.YAxis.MaxValue = 250;
                    //WOChart.YAxis.MinValue = 0;
                    WOChart.Locked = false;
                    WOChart.PlotArea.Border.Width = 0;
                    WOChart.YAxis.MinorTickMark = eAxisTickMark.None;
                    WOChart.DataLabel.ShowValue = true;
                    WOChart.DisplayBlanksAs = eDisplayBlanksAs.Gap;
                    WOChart.Series.Add(WODuration, MachinesNames);

                    //To remove grid lines
                    RemoveGridLines(ref WOChartExcel);
                }
                #endregion
            }
            catch (Exception e)
            {
                IntoFile("NoLogin Graph Error: " + e);
            }
            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "NoLoginReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "NoLoginReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
        }

        //UnAssignedWO Report Method PEStartTime(Login or WO Selection from DDL List) to Date(Submit) fields in table
        public async Task<string> UnAssignedWOReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\UnAssignedWOReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateSummarized = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "UnAssignedWOReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "UnAssignedWOReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetOA = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }
            if (worksheetOA == null)
            {
                try{
                worksheetOA = p.Workbook.Worksheets.Add("Summarized ", TemplateSummarized);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //Step1: Get the Summarized according to Name. for each WC from StartDate - EndDate
            //( Leave that many rows blank and Fill WorkCenter wise 1st.
           

            //DataTable for Consolidated Duration 
            DataTable DTNoLogin = new DataTable();
            DTNoLogin.Columns.Add("MachineID", typeof(int));
            DTNoLogin.Columns.Add("Duration", typeof(int));

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5;
            int Sno = 1;
            for (int i = 0; i < TotalDay + 1; i++)
            {
                int thisDatesDataStartsAt = Row;
                DateTime endDateTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = UsedDateForExcel.ToString("yyyy-MM-dd");

                DataTable HMIData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = null;
                    if (lowestLevel == "Plant")
                    {
                        //query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and !(PEStartTime is null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) )  ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and (PEStartTime is not null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) )  ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "Shop")
                    {
                        //query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and !(PEStartTime is null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and (PEStartTime is not null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "Cell")
                    {
                        //query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and !(PEStartTime is null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and (PEStartTime is not null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }
                    else if (lowestLevel == "WorkCentre")
                    {
                        //query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and !(PEStartTime is null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) group by MachineID order by PEStartTime; ";
                        query = "select  MachineID, PEStartTime,Date from i_facility_tal.dbo.tblhmiscreen where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and (PEStartTime is not null) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) group by MachineID order by PEStartTime; ";
                    }

                    //SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(HMIData);
                    mc.close();
                }

                double TodaysDuration = 0;
                for (int n = 0; n < HMIData.Rows.Count; n++)
                {
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        thisDatesDataStartsAt = Row;
                    }
                    int MachineID = Convert.ToInt32(HMIData.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = HierarchyData[3];

                    string Duration = null;
                    DateTime LoginTime = Convert.ToDateTime(HMIData.Rows[n][1]);
                    var newTimeSpan = new TimeSpan(05, 59, 59);
                    DateTime DayEndTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1) + newTimeSpan);

                    DateTime SubmitTime = DayEndTime;//Default Assign this Date
                    if (!string.IsNullOrEmpty(Convert.ToString(HMIData.Rows[n][2])))
                    {
                        SubmitTime = Convert.ToDateTime(HMIData.Rows[n][2]);
                    }

                    double DurInSeconds = 0;
                    string timeInSec = Convert.ToString(SubmitTime.Subtract(LoginTime).TotalSeconds);
                    double.TryParse(timeInSec, out DurInSeconds);
                    Duration = Math.Round(DurInSeconds / (60 * 60), 2).ToString();
                    TodaysDuration += DurInSeconds;

                    worksheet.Cells["G" + Row].Value = Convert.ToDateTime(LoginTime).ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells["H" + Row].Value = Convert.ToDateTime(SubmitTime).ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells["I" + Row].Value = Math.Round(Convert.ToDouble(Duration), 2);

                    DataRow dr = DTNoLogin.Select("MachineID = " + MachineID).FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                    if (dr != null)
                    {
                        Int64 DurationPrev = Convert.ToInt64(dr["Duration"]); //get lossduration and update it.
                        dr["Duration"] = (DurationPrev + DurInSeconds);
                    }
                    else
                    {
                        DTNoLogin.Rows.Add(MachineID, DurInSeconds);
                    }
                    //double Minutes = Math.Round(Convert.ToDateTime(HMIData.Rows[n][5]).Subtract(Convert.ToDateTime(HMIData.Rows[n][4])).TotalMinutes, 2);
                    Row++;
                }

                if (Row > thisDatesDataStartsAt)
                {

                    //thin border for all cells
                    var modelTable = worksheet.Cells["B" + thisDatesDataStartsAt + ":I" + Row];
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Insert OverAll NoLogin Time
                    worksheet.Cells["B" + Row + ":H" + Row].Merge = true;
                    worksheet.Cells["B" + Row].Value = "Duration";
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Font.Bold = true;
                    worksheet.Cells["I" + Row].Value = Math.Round(TodaysDuration / (60 * 60), 2);
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + (thisDatesDataStartsAt) + ":I" + "" + (Row - 1)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                }

                Row++;
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }
            worksheet.Cells["B" + 5 + ":I" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            //To Push Summarized Data
            Row = 6;
            Sno = 1;
            long SummarizedNoLoginDuration = 0;

            //to sort the NoLogin Datatable

            DataView dv = DTNoLogin.DefaultView;
            dv.Sort = "Duration desc";
            DataTable SortedDTNoLogin = dv.ToTable();

            for (int n = 0; n < SortedDTNoLogin.Rows.Count; n++)
            {
                int MachineID = Convert.ToInt32(SortedDTNoLogin.Rows[n][0]);
                List<string> HierarchyData = GetHierarchyData(MachineID);

                worksheetOA.Cells["B" + Row].Value = Sno++;
                worksheetOA.Cells["C" + Row].Value = HierarchyData[0];
                worksheetOA.Cells["D" + Row].Value = HierarchyData[1];
                worksheetOA.Cells["E" + Row].Value = HierarchyData[2];
                worksheetOA.Cells["F" + Row].Value = HierarchyData[3];

                worksheetOA.Cells["G" + Row].Value = Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd");
                worksheetOA.Cells["H" + Row].Value = Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd");

                string Duration = "0";
                DataRow dr = SortedDTNoLogin.Select("MachineID = " + MachineID).FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                if (dr != null)
                {
                    Duration = Convert.ToString(dr["Duration"]); //get lossduration and update it.
                }

                double DurDouble = 0;
                double.TryParse(Duration, out DurDouble);

                SummarizedNoLoginDuration += Convert.ToInt64(Duration);
                //Duration = TimeFromSeconds(Convert.ToInt64(Duration));
                // worksheetOA.Cells["I" + Row].Value = Duration;

                worksheetOA.Cells["Z" + Row].Value = Math.Round(DurDouble / 3600, 1); ;
                worksheetOA.Cells["I" + Row].Value = Math.Round(DurDouble / 3600, 1); ;

                //Making Text white
                //Color ColorHexWhite = System.Drawing.Color.White;
                //worksheetOA.Cells["Z6:Z" + Row + ""].Style.Font.Color.SetColor(ColorHexWhite);

                Row++;
            }

            //thin border for all cells
            var modelTable1 = worksheetOA.Cells["B5:I" + Row];
            // Assign borders
            modelTable1.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable1.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //worksheetOA.Cells["I" + Row].Value = TimeFromSeconds(Convert.ToInt64(SummarizedNoLoginDuration));
            double SumNoLogin1 = 0;
            SumNoLogin1 = Convert.ToDouble(SummarizedNoLoginDuration / 3600);
            double RoundOffSumNoLogin1 = Math.Round(SumNoLogin1, 1);
            worksheetOA.Cells["I" + Row].Value = RoundOffSumNoLogin1;

            worksheetOA.Cells["B" + Row + ":H" + Row].Merge = true;
            worksheetOA.Cells["B" + Row].Value = "Total Duration";
            worksheetOA.Cells["B" + Row + ":I" + Row].Style.Font.Bold = true;
            worksheetOA.Cells["B" + Row + ":I" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
            //Excel:: Border Around Cells.
            worksheetOA.Cells["B5:I" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);

            worksheet.View.ShowGridLines = false;
            worksheetOA.View.ShowGridLines = false;
            worksheetOA.Cells["B" + 5 + ":I" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheetOA.Cells[worksheetOA.Dimension.Address].AutoFitColumns();

            #region The Graph, Pareto Graph for UnassignedWO
            if (Row > 6)
            {
                Row = Row - 1;

                ExcelRange MachinesNames = worksheetOA.Cells["F6:F" + Row];
                ExcelRange WODuration = worksheetOA.Cells["Z6:Z" + Row];

                var WOChart = (ExcelBarChart)worksheetOA.Drawings.AddChart("Summarized" + System.DateTime.Now.ToString("dd-MM-yyyy"), eChartType.ColumnClustered);
                ExcelChart WOChartExcel = (ExcelChart)WOChart;
                #region To calculate the Auto width and Height for Graph position
                Row = Row + 5;
                int row = worksheetOA.Cells["A1:A" + Row].Rows;// cellAddress.Start.Row;
                int column = worksheetOA.Cells["I1:I" + Row].Columns;// cellAddress.Start.Column;

                int rowSpan = row;
                int colSpan = column;

                double height = 0, width = 0;
                for (int h = 0; h < rowSpan; h++)
                {
                    height += worksheetOA.Row(row + h).Height;
                }
                for (int w = 0; w < colSpan; w++)
                {
                    width += worksheetOA.Column(column + w).Width;
                }

                double pointToPixel = 0.75;

                height /= pointToPixel;
                width /= 0.1423;
                WOChart.SetPosition((int)height, (int)width);

                #endregion End Auto width and Height graph position

                WOChart.SetSize(600, 300);
                WOChart.Title.Text = "Unassigned WorkOrder";
                WOChart.Style = eChartStyle.Style15;
                WOChart.Legend.Position = eLegendPosition.Bottom;
                //WOChart.YAxis.MaxValue = 250;
                //WOChart.YAxis.MinValue = 0;
                WOChart.Locked = false;
                WOChart.PlotArea.Border.Width = 0;
                WOChart.YAxis.MinorTickMark = eAxisTickMark.None;
                WOChart.DataLabel.ShowValue = true;
                WOChart.DisplayBlanksAs = eDisplayBlanksAs.Gap;

                WOChart.Series.Add(WODuration, MachinesNames);

                //To remove grid lines
                RemoveGridLines(ref WOChartExcel);
            }
            #endregion

            // worksheetOA.Cells.AutoFitColumns(0);
            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "UnAssignedWOReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "UnAssignedWOReport " + lowestLevelName +" " + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
        }

        //MRR Report Method
        public async Task<string> MRRReportExcel(string StartDate, string EndDate, string PlantID, string PartsList = null, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            #region LowestLevel & Stuff
            //Step1: Get the Summarized according to Name. for each WC from StartDate - EndDate
            //( Leave that many rows blank and Fill WorkCenter wise 1st.
            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId && m.IsNormalWC == 0).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId && m.IsNormalWC == 0).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId && m.IsNormalWC == 0).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            #endregion

            #region Excel & Stuff
            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\MRRReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MRRReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MRRReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetGraph = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
            }
            catch { }

            if (worksheet == null)
            {
                try
                {
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }
            if (worksheetGraph == null)
            {
                try{
                worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            #endregion

            

            // List<string> PartNos = new List<string>();
            //PartNos = ExtractPartNo(PartsList);

            //string PartNos = ExtractPartNoString(PartsList);
            //required Format : '92200-02311-130|92209-02402-101|92308-02105-120(.*)'
            string PartNos = null;
            PartNos = ExtractPartNoStringRegex(PartsList);
            List<string> PartNoListString = PartsList.Split(',').ToList();

            //DataTable for Consolidated Data 
            DataTable DTUtil = new DataTable();
            DTUtil.Columns.Add("Plant", typeof(string));
            DTUtil.Columns.Add("Shop", typeof(string));
            DTUtil.Columns.Add("Cell", typeof(string));
            DTUtil.Columns.Add("WC", typeof(string));
            DTUtil.Columns.Add("TargetQty", typeof(double));
            DTUtil.Columns.Add("DeliveredQty", typeof(double));
            DTUtil.Columns.Add("RejectedQty", typeof(double));
            DTUtil.Columns.Add("ActualCuttingTime", typeof(double));
            DTUtil.Columns.Add("MaterialRemoved", typeof(double));
            DTUtil.Columns.Add("Date", typeof(string));
            DTUtil.Columns.Add("WONo", typeof(string));
            DTUtil.Columns.Add("OpNo", typeof(string));


            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5 + MacCount + 2;
            int Sno = 1;
            for (int i = 0; i < TotalDay + 1; i++)
            {
                double IndividualDateTarget = 0, IndividualDateDelivered = 0, IndividualDateRejected = 0, IndividualDateActualCuttingTime = 0, IndividualDateMaterialRemoved = 0, IndividualDateMR = 0, IndividualDateMRR = 0;

                DateTime endDateTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = UsedDateForExcel.ToString("yyyy-MM-dd");

                DataTable HMIData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = null;
                    //partno is optional so,
                    if (PartNos.Length == 0 || string.IsNullOrEmpty(PartNos))
                    {
                        //db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
                        if (lowestLevel == "Plant")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )   order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )   order by PEStartTime; ";
                        }
                        else if (lowestLevel == "Shop")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )   order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )   order by PEStartTime; ";
                        }
                        else if (lowestLevel == "Cell")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )  order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )  order by PEStartTime; ";
                        }
                        else if (lowestLevel == "WorkCentre")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )  order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )  order by PEStartTime; ";
                        }
                    }
                    else
                    {
                        if (lowestLevel == "Plant")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + " order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + " order by PEStartTime; ";
                        }
                        else if (lowestLevel == "Shop")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + "  order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + "  order by PEStartTime; ";
                        }
                        else if (lowestLevel == "Cell")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + "  order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + "  order by PEStartTime; ";
                        }
                        else if (lowestLevel == "WorkCentre")
                        {
                            //query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )   and PartNo REGEXP " + PartNos + "  order by PEStartTime; ";
                            query = "select * from i_facility_tal.dbo.tblhmiscreen where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ((IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )   and PartNo like " + PartNos + "  order by PEStartTime; ";
                        }
                    }

                    //SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(HMIData);
                    mc.close();
                }

                if (i == 0)
                {
                    //MacCount = (from DataRow dRow in HMIData.Rows
                    //            select new { col1 = dRow["MachineID"] }).Distinct().Count();

                    DataTable DistinctMacCount = new DataTable();
                    using (MsqlConnection mc = new MsqlConnection())
                    {
                        mc.open();
                        String query = null;
                        if (PartNos.Trim() == null)
                        {
                            //db.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
                            if (lowestLevel == "Plant")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1 ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )    order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1 ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )    order by PEStartTime; ";
                            }
                            else if (lowestLevel == "Shop")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )    order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )    order by PEStartTime; ";
                            }
                            else if (lowestLevel == "Cell")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )    order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )    order by PEStartTime; ";
                            }
                            else if (lowestLevel == "WorkCentre")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )   order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )   order by PEStartTime; ";
                            }
                        }
                        else
                        {
                            if (lowestLevel == "Plant")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + "  order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  ) and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  PlantId = " + plantId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + "  order by PEStartTime; ";
                            }
                            else if (lowestLevel == "Shop")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + "   order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  ShopID = " + shopId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + "   order by PEStartTime; ";
                            }
                            else if (lowestLevel == "Cell")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) and PartNo REGEXP " + PartNos + "   order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  CellID = " + cellId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) and PartNo like " + PartNos + "   order by PEStartTime; ";
                            }
                            else if (lowestLevel == "WorkCentre")
                            {
                                //query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) )   and PartNo REGEXP " + PartNos + "    order by PEStartTime; ";
                                query = "select count(distinct(MachineID)) from i_facility_tal.dbo.tblhmiscreen where CorrectedDate >= '" + frmDate.ToString("yyyy-MM-dd") + "' and  CorrectedDate <= '" + toDate.ToString("yyyy-MM-dd") + "' and isWorkOrder = 0 and ( isWorkInProgress = 1  )  and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where  IsNormalWC = 0 and  MachineID = " + wcId + "  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) )   and PartNo like " + PartNos + "    order by PEStartTime; ";
                            }
                        }

                        //SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                        SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                        da.Fill(DistinctMacCount);
                        mc.close();
                        MacCount = Convert.ToInt32(DistinctMacCount.Rows[0][0]);
                        Row = 5 + Convert.ToInt32(DistinctMacCount.Rows[0][0]) + 2 + 1;
                    }
                }

                for (int n = 0; n < HMIData.Rows.Count; n++)
                {
                    if (n == 0 && i != 0)
                    {
                        Row++;
                    }
                    int MachineID = Convert.ToInt32(HMIData.Rows[n][1]);

                    //Check if its MultiWo and Fetch Data.
                    int isMultiWo = Convert.ToInt32(HMIData.Rows[n][24]);

                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0]; //Plant Name
                    worksheet.Cells["D" + Row].Value = HierarchyData[1]; // Shop Name
                    worksheet.Cells["E" + Row].Value = HierarchyData[2]; //Cell Name
                    worksheet.Cells["F" + Row].Value = HierarchyData[3]; //WC Name
                    worksheet.Cells["G" + Row].Value = Convert.ToString(HMIData.Rows[n][14]); //Date

                    worksheet.Cells["H" + Row].Value = Convert.ToString(HMIData.Rows[n][10]); //WO No
                    worksheet.Cells["I" + Row].Value = Convert.ToString(HMIData.Rows[n][7]); //Part No
                    worksheet.Cells["J" + Row].Value = Convert.ToString(HMIData.Rows[n][8]); // Op No
                    worksheet.Cells["K" + Row].Value = Convert.ToString(HMIData.Rows[n][6]); // Projec
                    //worksheet.Cells["K" + Row].Value = Convert.ToInt32(HMIData.Rows[n][]);

                    string PartNo = Convert.ToString(HMIData.Rows[n][7]), WONo = Convert.ToString(HMIData.Rows[n][10]), OpNo = Convert.ToString(HMIData.Rows[n][8]);

                    double deliveredQty = 0;
                    string deliveredQtyString = Convert.ToString(HMIData.Rows[n][12]);
                    double.TryParse(deliveredQtyString, out deliveredQty);

                    string rejectedString = Convert.ToString(HMIData.Rows[n][9]);
                    double rejectedQty = 0;
                    double.TryParse(rejectedString, out rejectedQty);

                    worksheet.Cells["L" + Row].Value = Convert.ToString(HMIData.Rows[n][11]); // Target
                    worksheet.Cells["M" + Row].Value = deliveredQty; // Delivered
                    worksheet.Cells["N" + Row].Value = rejectedQty; // Rejected

                    //double ACuttingTime = GetActualCuttingTime(MachineID, PartNo, WONo, OpNo, UsedDateForExcel.ToString("yyyy-MM-dd"));
                    //2017-03-17

                    DateTime woStartTime = Convert.ToDateTime(HMIData.Rows[n][4]);
                    DateTime woEndTime = Convert.ToDateTime(HMIData.Rows[n][5]);
                    Task<double> ACuttingTime = GetGreen(UsedDateForExcel.ToString("yyyy-MM-dd"), woStartTime, woEndTime, MachineID);
                    worksheet.Cells["O" + Row].Value = Convert.ToDouble(ACuttingTime); // Actual Cutting Time


                    int hmiid = Convert.ToInt32(HMIData.Rows[n][0]);
                    int MainMultiWORow = Row;
                    double MRWeight = 0, MR = 0;
                    if (Convert.ToInt32(HMIData.Rows[n][24]) == 0) // is single WO
                    {
                        string MrQtyString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo).Select(m => m.MaterialRemovedQty).FirstOrDefault());
                        double.TryParse(MrQtyString, out MRWeight);
                        IndividualDateMR = MRWeight;
                        MR = MRWeight * (deliveredQty + rejectedQty);
                        worksheet.Cells["P" + Row].Value = MRWeight; // Material Removed
                        //2017-03-17
                        DTUtil.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3], Convert.ToString(HMIData.Rows[n][11]), deliveredQty, rejectedQty, Math.Round(Convert.ToDouble(ACuttingTime), 2), MR, Convert.ToString(HMIData.Rows[n][14]), Convert.ToString(HMIData.Rows[n][10]), Convert.ToString(HMIData.Rows[n][8]));

                    }
                    else
                    {
                        var MultiWOData = db.tbl_multiwoselection.Where(m => m.HMIID == hmiid).ToList();
                        foreach (var row in MultiWOData)
                        {

                            string PartNoInner = Convert.ToString(row.PartNo);
                            // if (PartsList.Contains(PartNoInner, StringComparer.OrdinalIgnoreCase))
                            //if (PartsList.Any(s => s.Equals(keyword, StringComparison.OrdinalIgnoreCase)))

                            if (PartNoListString.Any(s => s.Equals(PartNoInner, StringComparison.OrdinalIgnoreCase)))
                            {
                                Row++;
                                string WONoInner = Convert.ToString(row.WorkOrder), OpNoInner = Convert.ToString(row.OperationNo);
                                worksheet.Cells["H" + Row].Value = WONoInner; //WO No
                                worksheet.Cells["I" + Row].Value = Convert.ToString(row.PartNo); //Part No
                                worksheet.Cells["J" + Row].Value = OpNoInner; // Op No

                                int DeliveredQtyInner = 0, RejectedQtyInner = 0;
                                MRWeight = 0;
                                worksheet.Cells["L" + Row].Value = row.TargetQty; // Target
                                int.TryParse(Convert.ToString(row.DeliveredQty), out DeliveredQtyInner);
                                worksheet.Cells["M" + Row].Value = deliveredQty; // Delivered
                                int.TryParse(Convert.ToString(row.ScrapQty), out RejectedQtyInner);
                                worksheet.Cells["N" + Row].Value = rejectedQty; // Rejected
                                string MrQtyString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == PartNo && m.OpNo == OpNo).Select(m => m.MaterialRemovedQty).FirstOrDefault());
                                double.TryParse(MrQtyString, out MRWeight);
                                IndividualDateMR += MRWeight;
                                MR += MRWeight * (deliveredQty + rejectedQty);
                                worksheet.Cells["P" + Row].Value = MRWeight; // Material Removed

                                //2017-03-17
                                DTUtil.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3], row.TargetQty, DeliveredQtyInner, RejectedQtyInner, Math.Round(Convert.ToDouble(ACuttingTime), 2), MR, Convert.ToString(HMIData.Rows[n][14]), WONoInner, OpNoInner);
                            }
                        }
                    }

                    worksheet.Cells["Q" + MainMultiWORow].Value = MR; // Material Removed
                    double Target = Math.Round(Convert.ToDouble(HMIData.Rows[n][11]), 2);
                    //double Delivered = Math.Round(Convert.ToDouble(HMIData.Rows[n][12]), 2);
                    //double Rejected = Math.Round((Convert.ToDouble(HMIData.Rows[n][9])), 2);
                    double ActualCuttingTime = Math.Round(Convert.ToDouble(ACuttingTime), 2);
                    double MaterialRemoved = Math.Round(MR, 2);

                    //Row++;
                    //Sno++;
                    worksheet.Row(Row).CustomHeight = false;

                    IndividualDateTarget += Convert.ToInt32(HMIData.Rows[n][11]);
                    IndividualDateDelivered += deliveredQty;
                    IndividualDateRejected += rejectedQty;
                    IndividualDateActualCuttingTime += Convert.ToDouble(ACuttingTime);
                    IndividualDateMR += MRWeight;
                    IndividualDateMaterialRemoved += MR;
                }

                if (lowestLevel != "WorkCentre" && HMIData.Rows.Count > 0)
                {
                    worksheet.Cells["L" + Row].Value = IndividualDateTarget;
                    worksheet.Cells["M" + Row].Value = IndividualDateDelivered;
                    worksheet.Cells["N" + Row].Value = IndividualDateRejected;
                    worksheet.Cells["O" + Row].Value = IndividualDateActualCuttingTime;
                    worksheet.Cells["Q" + Row].Value = IndividualDateMR;
                    worksheet.Cells["P" + Row].Value = IndividualDateMaterialRemoved / IndividualDateDelivered;
                }
                if (HMIData.Rows.Count > 0)
                {
                    Row++;
                }
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }

            Sno = 1;
            Row = 5 + 1;
            //push Cummulative Data OVERAll for each Machine
            var MacCountDistinct = (from DataRow row in DTUtil.Rows
                                    select row["WC"]).Distinct();
            // MacCount = MacCountDistinct.Count();
            foreach (var MacRow in MacCountDistinct)
            {
                string MACName = Convert.ToString(MacRow);
                //Take Group of 
                worksheet.Cells["B" + Row].Value = Sno;
                worksheet.Cells["C" + Row].Value = (from DataRow row in DTUtil.Rows where row.Field<string>("WC") == @MACName select row["Plant"]).FirstOrDefault();
                worksheet.Cells["D" + Row].Value = (from DataRow row in DTUtil.Rows where row.Field<string>("WC") == @MACName select row["Shop"]).FirstOrDefault();
                worksheet.Cells["E" + Row].Value = (from DataRow row in DTUtil.Rows where row.Field<string>("WC") == @MACName select row["Cell"]).FirstOrDefault();
                worksheet.Cells["F" + Row].Value = MACName;

                worksheet.Cells["G" + Row + ":K" + Row].Merge = true;
                worksheet.Cells["G" + Row].Value = "Machine Wise Summarized for " + frmDate.ToString("yyyy-MM-dd") + "-" + toDate.ToString("yyyy-MM-dd");
                worksheet.Cells["G" + Row].Style.Font.Bold = true;

                var LVal = DTUtil.AsEnumerable().Where(r => r.Field<string>("WC") == MACName).Sum(x => x.Field<double>("TargetQty"));
                worksheet.Cells["L" + Row].Value = LVal;
                var MVal = DTUtil.AsEnumerable().Where(r => r.Field<string>("WC") == MACName).Sum(x => x.Field<double>("DeliveredQty"));
                worksheet.Cells["M" + Row].Value = MVal;
                var NVal = DTUtil.AsEnumerable().Where(r => r.Field<string>("WC") == MACName).Sum(x => x.Field<double>("RejectedQty"));
                worksheet.Cells["N" + Row].Value = NVal;
                var OVal = DTUtil.AsEnumerable().Where(r => r.Field<string>("WC") == MACName).Sum(x => x.Field<double>("ActualCuttingTime"));
                worksheet.Cells["O" + Row].Value = OVal;
                var PVal = DTUtil.AsEnumerable().Where(r => r.Field<string>("WC") == MACName).Sum(x => x.Field<double>("MaterialRemoved"));
                worksheet.Cells["P" + Row].Value = PVal;

                Sno++;
                Row++;
            }

            //Summarized (Selected)TopMostLevel
            Row = 5;
            double TargetQtySummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("TargetQty"));
            double DeliveredQtySummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("DeliveredQty"));
            double RejectedQtyTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("RejectedQty"));
            double ActualCuttingTimeTimeSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("ActualCuttingTime"));
            double MaterialRemovedSummation = DTUtil.AsEnumerable().Sum(x => x.Field<double>("MaterialRemoved"));

            worksheet.Cells["C" + Row + ":K" + Row].Merge = true;
            worksheet.Cells["C" + Row].Value = "Over All Summarized";
            worksheet.Cells["C" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C" + Row].Style.Font.Bold = true;
            worksheet.Cells["L" + Row].Value = TargetQtySummation;
            worksheet.Cells["M" + Row].Value = DeliveredQtySummation;
            worksheet.Cells["N" + Row].Value = RejectedQtyTimeSummation;
            worksheet.Cells["O" + Row].Value = ActualCuttingTimeTimeSummation;
            worksheet.Cells["P" + Row].Value = MaterialRemovedSummation;

            //worksheet.Cells["N" + Row].Value = Math.Round((valueAddingSummation / (24 * 60 * (TotalDay + 1))), 2);

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();


            #region Graphs

            var CorrectedDateDistinct = (from DataRow row in DTUtil.Rows
                                         select row["Date"]).Distinct();
            var WONoDistinct = (from DataRow row in DTUtil.Rows
                                select row["WONo"]).Distinct();
            var OpNoDistinct = (from DataRow row in DTUtil.Rows
                                select row["OpNo"]).Distinct();

            int rowG = 4, colG = 5;
            //Insert Header CorrectedDate & Op1,Op2....
            worksheetGraph.Cells["D" + rowG].Value = "Date";
            Dictionary<string, string> OpNameCol = new Dictionary<string, string>();


            foreach (var opG in OpNoDistinct)
            {
                string ColName = ExcelColumnFromNumber(colG);
                string OpName = Convert.ToString(opG);
                worksheetGraph.Cells[ColName + rowG].Value = "Op" + OpName;
                OpNameCol.Add(OpName, ColName);
                colG++;
            }
            rowG++;
            colG = 4;
            int rowPixel = 20, colPixel = 20, Opno = 0;
            int StartingRow = 0;
            foreach (var WoNoG in WONoDistinct)
            {
                StartingRow = rowG;
                string WONo = Convert.ToString(WoNoG);
                // worksheetGraph.Cells["C" + rowG].Value = WONo;
                foreach (var DateG in CorrectedDateDistinct)
                {
                    bool OpNoIsWritten = false;
                    string CorrectedDate = Convert.ToString(DateG);
                    foreach (var opG in OpNoDistinct)
                    {
                        //now get Value and if exists in DT for This WO and OpNo, get ColName Based on ColValue @ Row = 4.
                        string OpNo = Convert.ToString(opG);

                        double MRRToExcel = 0;
                        MRRToExcel = DTUtil.AsEnumerable().Where(x => x.Field<string>("WONo") == @WONo && x.Field<string>("OpNo") == @OpNo && x.Field<string>("Date") == @CorrectedDate).Sum(x => x.Field<double>("MaterialRemoved"));
                        if (MRRToExcel != 0 && OpNameCol.ContainsKey(OpNo))
                        {
                            worksheetGraph.Cells["D" + rowG].Value = CorrectedDate;
                            string ColName = OpNameCol[OpNo];
                            worksheetGraph.Cells[ColName + rowG].Value = MRRToExcel;
                            OpNoIsWritten = true;
                        }
                    }
                    if (OpNoIsWritten)
                    {
                        rowG++;
                    }
                }

                if (StartingRow < rowG)
                {
                    //The Graph goes here.
                    //to set Position of Graph in Excel
                    if (Opno != 0)
                    {
                        if (Opno % 3 == 0)
                        {
                            rowPixel += 360;
                            colPixel = 0;
                        }
                        else
                        {
                            colPixel += 410;
                        }
                    }
                    Opno++;

                    //How to increment Opno ?

                    ExcelChart chart01 = worksheetGraph.Drawings.AddChart("chart0" + WONo, eChartType.LineMarkers);
                    var chartOpNO = (ExcelLineChart)chart01;
                    chart01.SetPosition(rowPixel, colPixel);
                    chart01.Title.Font.Bold = true;
                    chart01.Title.Font.Size = 18;
                    chart01.YAxis.MinorTickMark = eAxisTickMark.None;
                    chart01.XAxis.MajorTickMark = eAxisTickMark.None;
                    //chart01.YAxis.MaxValue = (Convert.ToInt32(MaxCuttingTime) + 5);
                    chart01.YAxis.MinValue = 0;
                    chart01.Legend.Remove();
                    //chartOEE.YAxis.MaxValue = 100;
                    chart01.SetSize(400, 350);
                    chart01.Title.Text = "WorkOrder No.: " + WONo;
                    chart01.Style = eChartStyle.Style18;

                    //chart01.d
                    chartOpNO.DataLabel.ShowValue = true;

                    var DateRange = worksheetGraph.Cells["D" + StartingRow + ":D" + (rowG - 1)];
                    foreach (var opG in OpNoDistinct)
                    {
                        string OpNo = Convert.ToString(opG);
                        string ColName = OpNameCol[OpNo];
                        //var ran1 = worksheetGraph.Cells["A3:A10"];
                        var ran1 = worksheetGraph.Cells[ColName + "" + StartingRow + ":" + ColName + (rowG - 1)];
                        var serie1 = chart01.Series.Add(ran1, DateRange);
                    }

                    //var DateRange = worksheetGraph.Cells["D" + StartingRow + ":D" + (rowG - 1)];
                    //var serie1 = chart01.Series.Add(DateRange, worksheetGraph.Cells["D4"]);
                    //foreach (var opG in OpNoDistinct)
                    //{
                    //    string OpNo = Convert.ToString(opG);
                    //    string ColName = OpNameCol[OpNo];
                    //    var OpNoName = worksheetGraph.Cells[ColName+"4"];
                    //    var ran1 = worksheetGraph.Cells[ColName + "" + StartingRow + ":" + ColName + (rowG - 1)];
                    //    var serie2 = chart01.Series.Add(ran1, OpNoName);
                    //}

                    //chart01.Legend.Position = eLegendPosition.Bottom;
                }
            }
            #endregion

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "MRRReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "MRRReport " + lowestLevelName + " " + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
        }

        // Loss Analysis 2016-11-29 
        public async Task<string> LossAnalysisReportExcel(string StartDate, string EndDate, string ProdFAI, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;

            #region MacCount & LowestLevel
            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            #endregion

            #region Excel and Stuff

            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\LossDetailsReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "LossDetailsReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "LossDetailsReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetGraph = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
            }
            catch { }

            if (worksheet == null)
            {
                try
                {
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            } if (worksheetGraph == null)
            {
                try
                {
                    worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                }
                catch (Exception e)
                {
                }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            #endregion

            

            #region Get Machines List
            DataTable machin = new DataTable();
            DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime = frmDate.ToString("yyyy-MM-dd");
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = null;
                if (lowestLevel == "Plant")
                {
                    //query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + "  and IsNormalWC = 0  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + "  and IsNormalWC = 0  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "Shop")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "Cell")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "WorkCentre")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                //SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(machin);
                mc.close();
            }
            #endregion

            //DataTable for Consolidated Data 
            DataTable DTConsolidatedLosses = new DataTable();
            DTConsolidatedLosses.Columns.Add("Plant", typeof(string));
            DTConsolidatedLosses.Columns.Add("Shop", typeof(string));
            DTConsolidatedLosses.Columns.Add("Cell", typeof(string));
            DTConsolidatedLosses.Columns.Add("WCInvNo", typeof(string));
            DTConsolidatedLosses.Columns.Add("WCName", typeof(string));
            DTConsolidatedLosses.Columns.Add("CorrectedDate", typeof(string));

            //Get All Losses and Insert into DataTable
            DataTable LossCodesData = new DataTable();
            using (MsqlConnection mcLossCodes = new MsqlConnection())
            {
                mcLossCodes.open();
                string startDateTime1 = frmDate.ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);
                //string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "') end) ) and LossCodeID NOT IN (  "
                //            + "SELECT DISTINCT LossCodeID FROM (  "
                //            + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where LossCodesLevel1ID is not null and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) end) ) "
                //            + "UNION  "
                //            + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) end) )  "
                //            + ") AS derived ) order by LossCodesLevel1ID;";

                string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                            + "( (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "') ) ) and LossCodeID NOT IN (  "
                            + "SELECT DISTINCT LossCodeID FROM (  "
                            + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where LossCodesLevel1ID is not null and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                            + "(  (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) ) ) "
                            + "UNION  "
                            + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                            + "( (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) ) )  "
                            + ") AS derived ) order by LossCodesLevel1ID;";


                //SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                daLossCodesData.Fill(LossCodesData);
                mcLossCodes.close();
            }

            DTConsolidatedLosses.Columns.Add("MinorLoss", typeof(double));
            DTConsolidatedLosses.Columns["MinorLoss"].DefaultValue = "0";

            int LossesStartsATCol = 12;
            var LossesList = new List<KeyValuePair<int, string>>();

            #region LossCodes Into LossList
            for (int i = 0; i < LossCodesData.Rows.Count; i++)
            {
                int losscode = Convert.ToInt32(LossCodesData.Rows[i][0]);
                string losscodeName = Convert.ToString(LossCodesData.Rows[i][1]);

                var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == losscode).FirstOrDefault();
                int level = lossdata.LossCodesLevel;
                if (level == 3)
                {
                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                    var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                }

                else if (level == 2)
                {
                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                }
                else if (level == 1)
                {
                    if (losscode == 999)
                    {
                        losscodeName = "NoCode Entered";
                    }
                    else if (losscode == 9999)
                    {
                        losscodeName = "UnIdentified BreakDown";
                    }
                    else
                    {
                        losscodeName = lossdata.LossCode;
                    }
                }
                //losscodeName = LossHierarchy3rdLevel(losscode);
                DTConsolidatedLosses.Columns.Add(losscodeName, typeof(double));
                DTConsolidatedLosses.Columns[losscodeName].DefaultValue = "0";

                //Code to write LossesNames to Excel.
                string columnAlphabet = ExcelColumnFromNumber(LossesStartsATCol);

                worksheet.Cells[columnAlphabet + 4].Value = losscodeName;
                worksheet.Cells[columnAlphabet + 5].Value = "AF";

                LossesStartsATCol++;
                //Add the LossesToList
                LossesList.Add(new KeyValuePair<int, string>(losscode, losscodeName));
            }
            #endregion

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5 + machin.Rows.Count + 2; // Gap to Insert OverAll data. DataStartRow + MachinesCount + 2(1 for HighestLevel & another for Gap).
            int Sno = 1;
            string finalLossCol = null;


            for (int i = 0; i < TotalDay + 1; i++)
            {
                int StartingRowForToday = Row;
                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");

                int NumMacsToExcel = 0;
                for (int n = 0; n < machin.Rows.Count; n++)
                {
                    NumMacsToExcel++;
                    double CummulativeOfAllLosses = 0;
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        StartingRowForToday = Row;
                    }

                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    //worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = HierarchyData[4];
                    worksheet.Cells["G" + Row].Value = HierarchyData[3];

                    worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                    worksheet.Cells["I" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                    string CorrectedDateFormated = UsedDateForExcel.ToString("yyyy-MM-dd") + " 00:00:00";
                    DateTime StartDateFormatted = Convert.ToDateTime(UsedDateForExcel.ToString("yyyy-MM-dd") + " 00:00:00");

                    //Added this machineDetails into Datatable
                    string WCInvNoString = HierarchyData[3];
                    DataRow dr = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoString && r.Field<string>("CorrectedDate") == dateforMachine);
                    if (dr != null)
                    {
                        //do nothing
                    }
                    else
                    {
                        //plant, shop, cell, macINV, WcName, CorrectedDate, ValueAdding(Green/Operating), AvailableTime, SummationofSCTvsPP, Scrap,Rework,CuttingTime,DaysWorking, GodHours, TotalSTDHours, RejectionHours.
                        DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3], HierarchyData[4], dateforMachine);
                    }

                    //Now get & put Losses
                    // Push Loss Value into  DataTable & Excel
                    string correctedDate = UsedDateForExcel.ToString("yyyy-MM-dd");

                    #region Capture and Push Losses
                    int column = 12 + LossCodesData.Rows.Count - 1; // StartCol in Excel + TotalLosses
                    finalLossCol = ExcelColumnFromNumber(column);

                    //now push 0 for every other loss into excel
                    worksheet.Cells["L" + Row + ":" + finalLossCol + Row].Value = Convert.ToDouble(0.0);

                    if (ProdFAI == "OverAll")
                    {
                        double MinorLoss = 0;
                        string MLossString = Convert.ToString(db.tbloeedashboardvariables.Where(m => m.WCID == MachineID && m.StartDate == StartDateFormatted).Select(m => m.MinorLosses).FirstOrDefault());
                        double.TryParse(MLossString, out MinorLoss);
                        worksheet.Cells["K" + Row].Value = Math.Round(MinorLoss / 60, 2);

                        //to Capture and Push , Losses that occured.
                        List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSeconds(MachineID, correctedDate);
                        List<KeyValuePair<int, double>> BreakdowndurationList = GetAllBreakdownDurationSeconds(MachineID, correctedDate);
                        DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                        if (dr1 != null)
                        {
                            dr1["MinorLoss"] = Math.Round(MinorLoss / 60, 2);
                            CummulativeOfAllLosses += Math.Round(MinorLoss / 60, 2);

                            foreach (var loss in LossesdurationList)
                            {
                                int LossID = loss.Key;
                                double Duration = loss.Value;
                                var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossID).FirstOrDefault();
                                int level = lossdata.LossCodesLevel;
                                string losscodeName = null;

                                #region To Get LossCode Hierarchy
                                if (level == 3)
                                {
                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                                    var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                }
                                else if (level == 2)
                                {
                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                }
                                else if (level == 1)
                                {
                                    if (LossID == 999)
                                    {
                                        losscodeName = "NoCode Entered";
                                    }
                                    else
                                    {
                                        losscodeName = lossdata.LossCode;
                                    }
                                }
                                #endregion

                                //if (losscodeName == "Setup:Job")
                                //{
                                //}

                                int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 5);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                                double DurInHours = Convert.ToDouble(Math.Round((Duration / (60 * 60)), 1)); //To Hours:: 1 Decimal Place
                                worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                                dr1[losscodeName] = DurInHours;
                                CummulativeOfAllLosses += DurInHours;
                            }

                            foreach (var Breakdown in BreakdowndurationList)
                            {
                                int LossID = Breakdown.Key;
                                double Duration = Breakdown.Value;
                                var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossID).FirstOrDefault();
                                int level = lossdata.LossCodesLevel;
                                string losscodeName = null;

                                #region To Get LossCode Hierarchy
                                if (level == 3)
                                {
                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                                    var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                }
                                else if (level == 2)
                                {
                                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                }
                                else if (level == 1)
                                {
                                    if (LossID == 999)
                                    {
                                        losscodeName = "NoCode Entered";
                                    }
                                    else
                                    {
                                        losscodeName = lossdata.LossCode;
                                    }
                                }
                                #endregion

                                int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;

                                string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 5);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                                double DurInHours = Convert.ToDouble(Math.Round((Duration / (60 * 60)), 1)); //To Hours:: 1 Decimal Place
                                worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                                dr1[losscodeName] = DurInHours;
                                CummulativeOfAllLosses += DurInHours;
                            }
                        }
                    }
                    else
                    {
                        DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                        if (dr1 != null)
                        {
                            double MinorLossFinal = 0;
                            string MinorLossLossData = Convert.ToString(db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.Type == ProdFAI).Sum(m => m.MinorLoss));
                            double.TryParse(MinorLossLossData, out MinorLossFinal);
                            worksheet.Cells["K" + Row].Value = Math.Round(MinorLossFinal / 60, 2);
                            CummulativeOfAllLosses += Math.Round(MinorLossFinal / 60, 2);

                            ////Now get Loss rows data from tblwolosses and generate 1 single cummulative row of data
                            //var WODetails = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.Type == ProdFAI).ToList();
                            //foreach (var row in WODetails)
                            //{
                            //    int WOLossesHMIID = Convert.ToInt32(row.HMIID);
                            //    var LossesData = db.tblwolossesses.Where(m => m.HMIID == WOLossesHMIID).ToList();
                            //    foreach (var LossRow in LossesData)
                            //    {
                            //        double Duration = 0;
                            //        string LossName = LossRow.LossName;
                            //        DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                            //        if (dr1 != null)
                            //        {
                            //            string losscodeName = null;
                            //            int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                            //            string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 5);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                            //            double DurInHours = Convert.ToDouble(Math.Round((Duration / (60 * 60)), 1)); //To Hours:: 1 Decimal Place
                            //            worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                            //            dr1[losscodeName] = Convert.ToDouble(dr1[losscodeName]) + DurInHours;
                            //            CummulativeOfAllLosses += DurInHours;
                            //        }
                            //    }
                            //}

                            DataTable LossesDurationData = new DataTable();
                            using (MsqlConnection mcLossCodes = new MsqlConnection())
                            {
                                mcLossCodes.open();
                                string query = @"select sum(LossDuration),Level,LossName,LossCodeLevel1Name,LossCodeLevel2Name from i_facility_tal.dbo.tblwolossess Where HMIID in 
                                            ( SELECT HMIID FROM i_facility_tal.dbo.tblworeport where CorrectedDate = '" + correctedDate + "' and MachineID = '" + MachineID + "' and Type = '" + ProdFAI + "' ) group by LossName;";
                                SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                                daLossCodesData.Fill(LossesDurationData);
                                mcLossCodes.close();
                            }

                            for (int Lossloop = 0; Lossloop < LossesDurationData.Rows.Count; Lossloop++)
                            {
                                double Duration = Convert.ToDouble(LossesDurationData.Rows[Lossloop][0]);
                                int level = Convert.ToInt32(LossesDurationData.Rows[Lossloop][1]);
                                string losscodeName = null;

                                #region To Get LossCode Hierarchy
                                if (level == 3)
                                {
                                    string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][3]);
                                    string Level2Name = Convert.ToString(LossesDurationData.Rows[Lossloop][4]);
                                    string Level3Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                    losscodeName = Level1Name + " :: " + Level2Name + " : " + Level3Name;
                                }
                                else if (level == 2)
                                {
                                    string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][3]);
                                    string Level2Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                    losscodeName = Level1Name + ":" + Level2Name;
                                }
                                else if (level == 1)
                                {
                                    string Level1Name = Convert.ToString(LossesDurationData.Rows[Lossloop][2]);
                                    if (Level1Name == "999")
                                    {
                                        losscodeName = "NoCode Entered";
                                    }
                                    else
                                    {
                                        losscodeName = Level1Name;
                                    }
                                }
                                #endregion

                                int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 5);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                                double DurInHours = Convert.ToDouble(Math.Round((Duration / (60)), 1)); //To Hours:: 1 Decimal Place
                                worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInHours;
                                dr1[losscodeName] = DurInHours;
                                CummulativeOfAllLosses += DurInHours;
                            }
                        }
                    }
                    #endregion


                    worksheet.Cells["J" + Row].Value = Convert.ToDouble(Math.Round((CummulativeOfAllLosses), 1));
                    Row++;


                }//End of For Each Machine Loop

                //Stuff for entire day (of all WC's) Into DT
                DTConsolidatedLosses.Rows.Add("Summarized", "Summarized", "Summarized", "Summarized", "Summarized", dateforMachine);

                //Push each Date Cummulative. Loop through ExcelAddress and insert formula
                var rangeIndividualSummarized = worksheet.Cells["K4:" + finalLossCol + "4"];
                //rangeIndividualSummarized = worksheet.Cells["K4:CW4"];
                foreach (var rangeBase in rangeIndividualSummarized)
                {
                    string str = Convert.ToString(rangeBase);
                    string ExcelColAlphabet = Regex.Replace(str, "[^A-Z _]", "");
                    worksheet.Cells[ExcelColAlphabet + Row].Formula = "=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")";
                    //var a = worksheet.Cells[rangeBase.Address].Value;
                    var blah1 = worksheet.Calculate("=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")");

                    double LossVal = 0;
                    double.TryParse(Convert.ToString(blah1), out LossVal);
                    if (LossVal != 0.0)
                    {
                        string LossName = Convert.ToString(worksheet.Cells[ExcelColAlphabet + 4].Value);
                        DataRow dr = DTConsolidatedLosses.AsEnumerable().LastOrDefault(r => r.Field<string>("Plant") == "Summarized" && r.Field<string>("CorrectedDate") == dateforMachine);
                        if (dr != null)
                        {
                            dr[LossName] = LossVal;
                        }
                    }
                }

                //Total of Today into 
                //Insert Cummulative for today
                worksheet.Cells["C" + Row + ":G" + Row].Merge = true;
                worksheet.Cells["C" + Row].Value = "Summarized For";
                worksheet.Cells["H" + Row].Value = worksheet.Cells["H" + (Row - 1)].Value;
                worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Font.Bold = true;
                worksheet.Cells["J" + Row].Formula = "=SUM(J" + StartingRowForToday + ":J" + (Row - 1) + ")";

                //Cellwise Border for Today
                worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                //Excel:: Border Around Cells.
                worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                worksheet.Cells["B" + Row + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                Row++;

            } //End of All day's Loop




            #region OverAll Losses and Stuff
            Row = 5;
            Sno = 1;
            var WCInvNoList = (from DataRow row in DTConsolidatedLosses.Rows
                               where row["WCInvNo"] != "Summarized"
                               select row["WCInvNo"]).Distinct();

            foreach (var MacINV in WCInvNoList)
            {
                string WCInvNoStringOverAll = Convert.ToString(MacINV);
                DataRow drOverAll = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoStringOverAll);

                if (drOverAll != null)
                {
                    int MachineID = db.tblmachinedetails.Where(m => m.MachineInvNo == WCInvNoStringOverAll).Select(m => m.MachineID).SingleOrDefault();
                    string macDispName = db.tblmachinedetails.Where(m => m.MachineInvNo == WCInvNoStringOverAll).Select(m => m.MachineDispName).SingleOrDefault();
                    List<string> HierarchyData = GetHierarchyData(MachineID);
                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = macDispName;
                    worksheet.Cells["G" + Row].Value = HierarchyData[3];

                    worksheet.Cells["H" + Row].Value = (frmDate).ToString("yyyy-MM-dd");
                    worksheet.Cells["I" + Row].Value = (toDate).ToString("yyyy-MM-dd");


                    //Total of Losses
                    worksheet.Cells["J" + Row].Formula = "=SUM(K" + Row + ":" + finalLossCol + "" + Row + ")";

                    //OverAll Losses 
                    var range = worksheet.Cells["K4:" + finalLossCol + "" + 4];
                    int i = 11;

                    foreach (var rangeBase in range)
                    {
                        string LossNameVal = Convert.ToString(rangeBase.Value);
                        string LossName = Convert.ToString(worksheet.Cells[rangeBase.Address].Value);
                        double LossValToExcel = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @WCInvNoStringOverAll).Sum(x => x.Field<double>(@LossNameVal));
                        string ColumnForThisLoss = ExcelColumnFromNumber(i++);
                        worksheet.Cells[ColumnForThisLoss + "" + Row].Value = Math.Round(LossValToExcel, 2);
                    }
                } //End of if(drOverAll != null)
                Row++;
            }

            // Borders and Stuff for Cummulative Data.
            //Cellwise Border for Today
            worksheet.Cells["B4:" + finalLossCol + "" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["B4:" + finalLossCol + "" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["B4:" + finalLossCol + "" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells["B4:" + finalLossCol + "" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            //Excel:: Border Around Cells.
            worksheet.Cells["B5:" + finalLossCol + "" + (Row)].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
            worksheet.Cells["B" + Row + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);


            //Total of Losses OverAll, Summarized of Total per Day
            worksheet.Cells["J" + Row].Formula = "=SUM(J5:J" + (Row - 1) + ")";

            //Cummulative Losses into DT and Occured Losses into List and Identified and UnIdentified Losses
            var rangeFinalLosses = worksheet.Cells["K4:" + finalLossCol + "4"];
            List<KeyValuePair<string, double>> AllOccuredLosses = new List<KeyValuePair<string, double>>();
            int j = 11;
            double IdentifiedLoss = 0;
            double UnIdentifiedLoss = 0;
            foreach (var rangeBase in rangeFinalLosses)
            {
                string LossName = Convert.ToString(rangeBase.Value);
                string LossNameAddress = Convert.ToString(rangeBase.Address);
                if (LossName == "MinorLoss")
                {

                }
                else
                {
                    if (string.IsNullOrEmpty(LossName))
                    {
                    }
                    double thisLossValue = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("Plant") != "Summarized").Sum(x => x.Field<double>(@LossName));
                    string ColumnForThisLoss = ExcelColumnFromNumber(j++);
                    worksheet.Cells[ColumnForThisLoss + "" + Row].Formula = "=SUM(" + ColumnForThisLoss + 5 + ":" + ColumnForThisLoss + "" + (Row - 1) + ")";
                    if (thisLossValue > 0)
                    {
                        if (LossName == "NoCode Entered" || LossName == "Unidentified Breakdown")
                        {
                            UnIdentifiedLoss += thisLossValue;
                        }
                        else
                        {
                            IdentifiedLoss += thisLossValue;
                        }
                        AllOccuredLosses.Add(new KeyValuePair<string, double>(LossNameAddress, Math.Round(thisLossValue, 1)));
                    }
                }
            }

            #endregion

            #region GRAPHS
            //Create the chart
            List<double> ForAvg = new List<double>();
            if (machin.Rows.Count > 0)
            {

                #region lOSSES TOP 5 GRAPH
                //1. Get Top 5 Losses ColName in excel.
                //2. Generate the comma seperated String format .
                //sort the list
                AllOccuredLosses.Sort(Compare2);
                AllOccuredLosses = AllOccuredLosses.OrderByDescending(x => x.Value).ToList();

                #region Percentage Data into Graph sheet.

                var SumOfAllLosses = AllOccuredLosses.Sum(x => x.Value);
                j = 3;
                int CellRow = 5;
                foreach (var item in AllOccuredLosses)
                {
                    string LossNameCell = item.Key;
                    double LossValue = item.Value;
                    string ColumnForThisLoss = ExcelColumnFromNumber(j++);
                    worksheetGraph.Cells[ColumnForThisLoss + CellRow].Value = worksheet.Cells[LossNameCell].Value;
                    worksheetGraph.Cells[ColumnForThisLoss + (CellRow + 1)].Value = LossValue;
                    double InPercentage = Math.Round(((LossValue / SumOfAllLosses) * 100), 2);
                    worksheetGraph.Cells[ColumnForThisLoss + (CellRow + 2)].Value = InPercentage + "%";
                }

                //Cellwise Border
                string ColumnForEndOfLoss = ExcelColumnFromNumber(AllOccuredLosses.Count + 2);
                worksheetGraph.Cells["B5:" + ColumnForEndOfLoss + "7"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                worksheetGraph.Cells["B5:" + ColumnForEndOfLoss + "7"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                worksheetGraph.Cells["B5:" + ColumnForEndOfLoss + "7"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                worksheetGraph.Cells["B5:" + ColumnForEndOfLoss + "7"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                #endregion

                //Now construct string from top 5 losses.

                int LooperTop5 = 0;
                string CellsOfTop5LossColNames = null;
                string CellsOfTop5LossColValues = null;
                bool isNoCodeATopper = false;
                double NonOthersTotalInHours = 0;
                foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                {
                    if (LooperTop5 < 5)
                    {
                        string a = loss.Key;
                        double b = loss.Value;
                        var outputJustColName = Regex.Replace(a, @"[\d-]", string.Empty);
                        //string LossCol = Convert.ToString(outputJustColName);
                        string LossCol = a;
                        string lossName = Convert.ToString(worksheet.Cells[a].Value);
                        if (lossName != "NoCode Entered")
                        {
                            if (LooperTop5 == 0)
                            {
                                CellsOfTop5LossColNames = LossCol;
                                CellsOfTop5LossColValues = outputJustColName + Row;
                            }
                            else
                            {
                                CellsOfTop5LossColNames += "," + LossCol;
                                CellsOfTop5LossColValues += "," + outputJustColName + Row;
                            }
                            NonOthersTotalInHours += b;
                            LooperTop5++;
                        }
                        else
                        {
                            isNoCodeATopper = true;
                            NonOthersTotalInHours += b;
                        }
                    }
                    else
                    {
                        break;
                    }

                }

                //Calculate Others Time && if necessary remove NoCode Time
                double OthersTotalInHours = 0;
                if (isNoCodeATopper)
                {
                    OthersTotalInHours = SumOfAllLosses - NonOthersTotalInHours;
                }
                else
                {
                    foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                    {
                        string a = loss.Key;
                        string LossCol = a;
                        double b = loss.Value;
                        string lossName = Convert.ToString(worksheet.Cells[a].Value);
                        if (lossName == "NoCode Entered")
                        {
                            OthersTotalInHours = SumOfAllLosses - NonOthersTotalInHours - b;
                        }
                    }
                }

                //Now Append "Others" to graph
                int column = 20 + LossCodesData.Rows.Count; // StartCol in Excel + TotalLosses
                finalLossCol = ExcelColumnFromNumber(column);

                worksheet.Cells[finalLossCol + "1"].Value = "Others";
                worksheet.Cells[finalLossCol + "2"].Value = Math.Round(OthersTotalInHours, 1);

                if (CellsOfTop5LossColNames == null)
                {
                    var akl = worksheetGraph.Cells[10, 3].Address;
                    CellsOfTop5LossColNames = worksheet.Cells[finalLossCol + "1"].Address;
                    CellsOfTop5LossColValues = worksheet.Cells[finalLossCol + "2"].Address;
                }
                else
                {
                    var akl = worksheet.Cells[finalLossCol + "1"].Address;
                    CellsOfTop5LossColNames += "," + worksheet.Cells[finalLossCol + "1"].Address;
                    CellsOfTop5LossColValues += "," + worksheet.Cells[finalLossCol + "2"].Address;
                }

                ExcelChart chartTop5Losses1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartTop5Losses", eChartType.ColumnClustered);
                var chartTop5Losses = (ExcelBarChart)chartTop5Losses1;
                chartTop5Losses.SetSize(500, 400); //Width,Height
                chartTop5Losses.SetPosition(140, 10); //PixelTop,Pixelleft
                //string blah = "CY11,CZ11,DA11,DC11,DD11"; //This Works 
                //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells[blah];
                //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells["CY11,CZ11,DA11,DC11,DD11"];
                //ExcelRange erLossesRangechartTop5LossesNAMES = worksheet.Cells["CY3,CZ3,DA3,DC3,DD3"];

                ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfTop5LossColValues];
                ExcelRange erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfTop5LossColNames];
                chartTop5Losses.Title.Text = "LOSSES  ";
                chartTop5Losses.Style = eChartStyle.Style19;
                chartTop5Losses.Legend.Remove();
                chartTop5Losses.DataLabel.ShowValue = true;
                //chartTop5Losses.DataLabel.Font.Size = 8;
                //chartTop5Losses.Legend.Font.Size = 8;
                chartTop5Losses.XAxis.Font.Size = 8;
                chartTop5Losses.YAxis.MinorTickMark = eAxisTickMark.None;
                chartTop5Losses.XAxis.MajorTickMark = eAxisTickMark.None;
                chartTop5Losses.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                RemoveGridLines(ref chartTop5Losses1);

                #endregion

                #region Identified & UnIdentified Losses "scary"
                worksheetGraph.Cells["A1"].Value = "Ratio of Losses";
                //worksheetGraph.Cells["A2"].Value = "UnIdentifiedLoss";

                double IdentifiedLossPercentage = (IdentifiedLoss / (IdentifiedLoss + UnIdentifiedLoss)) * 100;
                double UnIdentifiedLossPercentage = (UnIdentifiedLoss / (IdentifiedLoss + UnIdentifiedLoss)) * 100;
                worksheetGraph.Cells["B1"].Value = Math.Round(IdentifiedLossPercentage, 0);
                worksheetGraph.Cells["B2"].Value = Math.Round(UnIdentifiedLossPercentage, 0);

                erLossesRangechartTop5LossesvALUE = worksheetGraph.Cells["B1"];
                ExcelRange erLossesRangechartTop5LossesvALUE1 = worksheetGraph.Cells["B2"];
                erLossesRangechartTop5LossesNAMES = worksheetGraph.Cells["A1"];

                ExcelChart chartIDAndUnID1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("TypesOfLosses", eChartType.ColumnStacked);
                var chartIDAndUnID = (ExcelBarChart)chartIDAndUnID1;
                chartIDAndUnID.SetSize(500, 400);
                chartIDAndUnID.SetPosition(140, 520);

                chartIDAndUnID.Title.Text = "Identified Losses  ";
                chartIDAndUnID.Style = eChartStyle.Style18;
                chartIDAndUnID.Legend.Position = eLegendPosition.Bottom;
                //chartIDAndUnID.Legend.Remove();
                chartIDAndUnID.YAxis.MaxValue = 100;
                chartIDAndUnID.YAxis.MinValue = 0;
                chartIDAndUnID.Locked = false;
                chartIDAndUnID.PlotArea.Border.Width = 0;
                chartIDAndUnID.YAxis.MinorTickMark = eAxisTickMark.None;
                chartIDAndUnID.DataLabel.ShowValue = true;
                //chartAllLosses.DataLabel.ShowValue = true;
                var thisYearSeries = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES));
                thisYearSeries.Header = "Identified Losses";
                var lastYearSeries = (ExcelChartSerie)(chartIDAndUnID.Series.Add(erLossesRangechartTop5LossesvALUE1, erLossesRangechartTop5LossesNAMES));
                lastYearSeries.Header = "UnIdentified Losses";
                RemoveGridLines(ref chartIDAndUnID1);

                #region OLD
                ////////////////////
                //have to remove cat nodes from each series so excel autonums 1 and 2 in xaxis
                //var chartXml = chartIDAndUnID.ChartXml;
                //var nsm = new XmlNamespaceManager(chartXml.NameTable);

                //var nsuri = chartXml.DocumentElement.NamespaceURI;
                //nsm.AddNamespace("c", nsuri);

                ////Get the Series ref and its cat
                //var serNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:bar3DChart/c:ser", nsm);
                //foreach (XmlNode serNode in serNodes)
                //{
                //    //Cell any cell reference and replace it with a string literal list
                //    var catNode = serNode.SelectSingleNode("c:cat", nsm);
                //    catNode.RemoveAll();

                //    //Create the string list elements
                //    var ptCountNode = chartXml.CreateElement("c:ptCount", nsuri);
                //    ptCountNode.Attributes.Append(chartXml.CreateAttribute("val", nsuri));
                //    ptCountNode.Attributes[0].Value = "2";

                //    var v0Node = chartXml.CreateElement("c:v", nsuri);
                //    v0Node.InnerText = "opening";
                //    var pt0Node = chartXml.CreateElement("c:pt", nsuri);
                //    pt0Node.AppendChild(v0Node);
                //    pt0Node.Attributes.Append(chartXml.CreateAttribute("idx", nsuri));
                //    pt0Node.Attributes[0].Value = "0";

                //    var v1Node = chartXml.CreateElement("c:v", nsuri);
                //    v1Node.InnerText = "closing";
                //    var pt1Node = chartXml.CreateElement("c:pt", nsuri);
                //    pt1Node.AppendChild(v1Node);
                //    pt1Node.Attributes.Append(chartXml.CreateAttribute("idx", nsuri));
                //    pt1Node.Attributes[0].Value = "1";

                //    //Create the string list node
                //    var strLitNode = chartXml.CreateElement("c:strLit", nsuri);
                //    strLitNode.AppendChild(ptCountNode);
                //    strLitNode.AppendChild(pt0Node);
                //    strLitNode.AppendChild(pt1Node);
                //    catNode.AppendChild(strLitNode);
                //}
                //pck.Save();
                #endregion
                #region Experiment to Send Data to Excel Chart in Template
                //OfficeOpenXml.FormulaParsing.Excel.Application xlApp;
                //Excel.Workbook xlWorkBook;
                //Excel.Worksheet xlWorkSheet;
                //object misValue = System.Reflection.Missing.Value;

                //Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
                //Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
                //Excel.Chart chartPage = myChart.Chart;

                //Excel.Range chartRange;
                //chartRange = xlWorkSheet.get_Range("A1", "d5");
                //chartPage.SetSourceData(chartRange, misValue);
                //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

                //xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                //xlWorkBook.Close(true, misValue, misValue);
                //xlApp.Quit();

                //var piechart = worksheet.Drawings["A"] as ExcelBarChart;
                //piechart.Style = eChartStyle.Style26;
                //chartIDAndUnID.Style = eChartStyle.Style26;

                //////////////////////////////


                #endregion

                #endregion End of Identified & UnIdentified Losses

                #region All Losses Chart

                CellsOfTop5LossColNames = null;
                CellsOfTop5LossColValues = null;
                int Looper = 0;
                foreach (KeyValuePair<string, double> loss in AllOccuredLosses)
                {
                    string a = loss.Key;
                    double b = loss.Value;

                    var outputJustColName = Regex.Replace(a, @"[\d-]", string.Empty);
                    //string LossCol = Convert.ToString(outputJustColName);
                    string LossCol = a;
                    if (a == "0")
                    {
                    }
                    if (Looper == 0)
                    {
                        CellsOfTop5LossColNames = LossCol;
                        CellsOfTop5LossColValues = outputJustColName + Row;
                    }
                    else
                    {
                        CellsOfTop5LossColNames += "," + LossCol;
                        CellsOfTop5LossColValues += "," + outputJustColName + Row;
                    }
                    Looper++;
                }

                if (CellsOfTop5LossColValues != null && CellsOfTop5LossColNames != null)
                {
                    ExcelChart chartAllLosses1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartAllLosses", eChartType.ColumnClustered);
                    var chartAllLosses = (ExcelBarChart)chartAllLosses1;
                    chartAllLosses.SetSize(1200, 500);
                    chartAllLosses.SetPosition(550, 10);
                    erLossesRangechartTop5LossesvALUE = worksheet.Cells[CellsOfTop5LossColValues];
                    erLossesRangechartTop5LossesNAMES = worksheet.Cells[CellsOfTop5LossColNames];
                    chartAllLosses.Title.Text = "All LOSSES ";
                    chartAllLosses.Style = eChartStyle.Style25;
                    chartAllLosses.Legend.Remove();
                    chartAllLosses.DataLabel.ShowValue = true;
                    //chartAllLosses.DataLabel.Font.Size = 8;
                    //chartAllLosses.Legend.Font.Size = 8;
                    chartAllLosses.YAxis.MinorTickMark = eAxisTickMark.None;
                    chartAllLosses.XAxis.MajorTickMark = eAxisTickMark.None;
                    chartAllLosses.XAxis.Font.Size = 8;
                    chartAllLosses.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);

                    //Get reference of Graph to Remove GridLines
                    RemoveGridLines(ref chartAllLosses1);
                }

                #endregion

                #region  Losses Trend :: All 5 Topper's
                var queryLinq = (from cell in worksheet.Cells["C:C"]
                                 where cell.Value is string && (string)cell.Value == "Summarized For"
                                 select cell);
                int LossesLooper = 1;
                int PositionY = 1060;
                int PositionX = 10;

                int GraphNo = 0;
                foreach (var Loss in AllOccuredLosses)
                {
                    if (LossesLooper <= 5)
                    {
                        ForAvg.Clear();
                        string a = Loss.Key;
                        double b = Loss.Value;
                        var outputJustColName = Regex.Replace(a, @"[\d-]", string.Empty);
                        //string LossCol = Convert.ToString(outputJustColName);
                        string LossCol = a;
                        string lossName = Convert.ToString(worksheet.Cells[a].Value);
                        if (lossName != "NoCode Entered")
                        {
                            string CellsOfOEEYAxis = null;
                            string CellsOfOEEXAxis = null;
                            string CellColString = Convert.ToString(Loss.Key);
                            string LossName = Convert.ToString(worksheet.Cells[CellColString].Value);
                            foreach (var cell in queryLinq)
                            {
                                string CellRowString = cell.Address;
                                string RowNum = Regex.Replace(CellRowString, "[^0-9 _]", string.Empty);
                                string ColName = Regex.Replace(CellColString, "[0-9 _]", string.Empty);

                                if (CellsOfOEEXAxis == null)
                                {
                                    CellsOfOEEXAxis = "H" + (Convert.ToInt32(RowNum) - 1);
                                }
                                else
                                {
                                    CellsOfOEEXAxis += ",H" + (Convert.ToInt32(RowNum) - 1);
                                }

                                if (CellsOfOEEYAxis == null)
                                {
                                    string colPlusrow = ColName + RowNum;
                                    CellsOfOEEYAxis = colPlusrow;
                                    double CellVal = 0;
                                    string CellValString = worksheet.Calculate(worksheet.Cells[colPlusrow].Formula).ToString();
                                    if (double.TryParse(CellValString, out CellVal))
                                    {
                                        ForAvg.Add(CellVal);
                                    }
                                }
                                else
                                {
                                    string colPlusrow = ColName + RowNum;
                                    CellsOfOEEYAxis += "," + colPlusrow;
                                    double CellVal = 0;

                                    string CellValString = worksheet.Calculate(worksheet.Cells[colPlusrow].Formula).ToString();
                                    if (double.TryParse(CellValString, out CellVal))
                                    {
                                        ForAvg.Add(CellVal);
                                    }
                                }
                            }

                            ExcelRange erTopLossesTrendYValues = worksheet.Cells[CellsOfOEEYAxis];
                            ExcelRange erTopLossesTrendXNames = worksheet.Cells[CellsOfOEEXAxis];
                            ExcelChart chartLossses1Trend1 = (ExcelLineChart)worksheetGraph.Drawings.AddChart("TrendChartTop" + LossesLooper, eChartType.LineMarkers);

                            GraphNo++;
                            if (GraphNo == 1)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["A" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["B" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["B50:B" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["A50:A" + (50 + ForAvg.Count - 1)]);
                            }
                            if (GraphNo == 2)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["C" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["D" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["D50:D" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["C50:C" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 3)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["E" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["F" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["F50:F" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["E50:E" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 4)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["G" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["H" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["H50:H" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["G50:G" + (50 + ForAvg.Count - 1)]);

                            }
                            if (GraphNo == 5)
                            {
                                for (int i = 0; i < ForAvg.Count; i++)
                                {
                                    worksheetGraph.Cells["I" + (50 + i)].Value = "Avg";
                                    worksheetGraph.Cells["J" + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
                                }
                                var chartTypeL1 = chartLossses1Trend1.PlotArea.ChartTypes.Add(eChartType.Line);
                                var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells["J50:J" + (50 + ForAvg.Count - 1)], worksheetGraph.Cells["I50:I" + (50 + ForAvg.Count - 1)]);

                            }

                            if (erTopLossesTrendXNames != null && erTopLossesTrendYValues != null)
                            {
                                var chartLossses1Trend = (ExcelLineChart)chartLossses1Trend1;
                                chartLossses1Trend.SetSize(300, 300);
                                chartLossses1Trend.SetPosition(PositionY, 10 + (((LossesLooper - 1)) * 300));
                                //chartLossses1Trend.Title.Text = "Top" + LossesLooper + " Trend Chart ";
                                chartLossses1Trend.Title.Text = LossName;
                                chartLossses1Trend.Style = eChartStyle.Style8;
                                chartLossses1Trend.Legend.Remove();
                                //chartLossses1Trend.YAxis.MaxValue = 100;
                                chartLossses1Trend.DataLabel.ShowValue = true;
                                //chartLossses1Trend.DataLabel.Font.Size = 6.0F;
                                chartLossses1Trend.PlotArea.Border.Width = 0;
                                chartLossses1Trend.YAxis.MinorTickMark = eAxisTickMark.None;
                                //chartLossses1Trend.YAxis.MajorTickMark = eAxisTickMark.None;
                                //chartLossses1Trend.XAxis.MinorTickMark = eAxisTickMark.None;
                                chartLossses1Trend.XAxis.MajorTickMark = eAxisTickMark.None;
                                //chartLossses1Trend.XAxis.MinorTickMark = eAxisTickMark.None;
                                chartLossses1Trend.Series.Add(erTopLossesTrendYValues, erTopLossesTrendXNames);

                                //Get reference of Graph to Remove GridLines
                                RemoveGridLines(ref chartLossses1Trend1);
                            }
                            #region OLD
                            //chartXml = chartLossses1Trend.ChartXml;
                            // nsuri = chartXml.DocumentElement.NamespaceURI;
                            // nsm = new XmlNamespaceManager(chartXml.NameTable);
                            //nsm.AddNamespace("c", nsuri);

                            ////XY Scatter plots have 2 value axis and no category
                            //valAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:valAx", nsm);
                            //if (valAxisNodes != null && valAxisNodes.Count > 0)
                            //    foreach (XmlNode valAxisNode in valAxisNodes)
                            //    {
                            //        var major = valAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                            //        if (major != null)
                            //            valAxisNode.RemoveChild(major);

                            //        var minor = valAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                            //        if (minor != null)
                            //            valAxisNode.RemoveChild(minor);
                            //    }

                            ////Other charts can have a category axis
                            //catAxisNodes = chartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea/c:catAx", nsm);
                            //if (catAxisNodes != null && catAxisNodes.Count > 0)
                            //    foreach (XmlNode catAxisNode in catAxisNodes)
                            //    {
                            //        var major = catAxisNode.SelectSingleNode("c:majorGridlines", nsm);
                            //        if (major != null)
                            //            catAxisNode.RemoveChild(major);

                            //        var minor = catAxisNode.SelectSingleNode("c:minorGridlines", nsm);
                            //        if (minor != null)
                            //            catAxisNode.RemoveChild(minor);
                            //    }
                            #endregion

                            LossesLooper++;
                        }
                    }
                    else
                    {
                        break;
                    }
                }

                #endregion

            }
            #endregion

            //Apply style to Losses header
            int col = 12 + LossCodesData.Rows.Count - 1; // StartCol in Excel + TotalLosses
            finalLossCol = ExcelColumnFromNumber(col);
            //worksheetGraph.Cells[worksheet.Dimension.Address].AutoFitColumns();

            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#32CD32");//#32CD32:lightgreen //B8C9E9
            worksheet.Cells["K4:" + finalLossCol + "" + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["K4:" + finalLossCol + "" + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);
            worksheet.Cells["K4:" + finalLossCol + "" + 4].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            worksheet.Cells["K4:" + finalLossCol + "" + 4].Style.WrapText = true;
            worksheetGraph.Cells["A5:" + finalLossCol + "" + 5].Style.WrapText = true;

            //worksheetGraph.Cells["A1:B2"].Style.Font.Color.SetColor(Color.White);
            worksheet.Row(4).Height = 70;
            worksheetGraph.Row(5).Height = 90;
            worksheet.View.ShowGridLines = false;
            worksheetGraph.View.ShowGridLines = false;
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheetGraph.Cells[worksheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //worksheetGraph.Cells[worksheet.Dimension.Address].AutoFitColumns();

            #region Save and Download
            p.Save();

            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "LossDetailsReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "LossDetailsReport " + lowestLevelName + " " + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
            #endregion

        }

        // JOB or Operations Report 2016-12-04
        public async Task<string> JOBReportExcel(string StartDate, string EndDate, string ProdFAI, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null, string Operator = null)
        {
            string lowestLevelName = null;
            #region Excel and Stuff

            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);

            DateTime toDate = Convert.ToDateTime(EndDate);
            EndDate = Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 00:00:00");
            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\JobReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            #region MacCount & LowestLevel
            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            #endregion

            lowestLevelName = lowestLevelName.Trim();

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "JobReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "JobReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetGraph = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }
            if (worksheetGraph == null)
            {
                try {
                    worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                }
                catch (Exception e)
                { }
            }
            if (worksheet == null)
            {
                IntoFile(" JobReport :: WorkSheet Not Created.");
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            #endregion

            

            #region Get Machines List
            DataTable machin = new DataTable();
            DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime = frmDate.ToString("yyyy-MM-dd");
            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String query1 = null;
            if (lowestLevel == "Plant")
            {
                //query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and ManualWCID IS NULL  AND((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and ManualWCID IS NULL  AND((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "Shop")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + " and ManualWCID IS NULL and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + " and ManualWCID IS NULL and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "Cell")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + " and ManualWCID IS NULL and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + " and ManualWCID IS NULL and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "WorkCentre")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + " and ManualWCID IS NULL and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + " and ManualWCID IS NULL and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
            da1.Fill(machin);
            mc.close();
            #endregion

            //DataTable for Consolidated Data 
            DataTable DTConsolidatedLosses = new DataTable();
            DTConsolidatedLosses.Columns.Add("Plant", typeof(string));
            DTConsolidatedLosses.Columns.Add("Shop", typeof(string));
            DTConsolidatedLosses.Columns.Add("Cell", typeof(string));
            DTConsolidatedLosses.Columns.Add("WCInvNo", typeof(string));
            DTConsolidatedLosses.Columns.Add("WCName", typeof(string));
            DTConsolidatedLosses.Columns.Add("CorrectedDate", typeof(string));
            //DTConsolidatedLosses.Columns.Add("HMIID", typeof(string));
            DTConsolidatedLosses.Columns.Add("OpName", typeof(string));
            DTConsolidatedLosses.Columns.Add("WOPF", typeof(int));
            DTConsolidatedLosses.Columns.Add("WOProcessed", typeof(int));
            DTConsolidatedLosses.Columns.Add("TotalWOQty", typeof(int));
            DTConsolidatedLosses.Columns.Add("TotalTarget", typeof(int));
            DTConsolidatedLosses.Columns.Add("TotalDelivered", typeof(int));
            DTConsolidatedLosses.Columns.Add("TargetNC", typeof(double));
            DTConsolidatedLosses.Columns.Add("TotalValueAdding", typeof(double));
            DTConsolidatedLosses.Columns.Add("TotalLosses", typeof(double));
            DTConsolidatedLosses.Columns.Add("TotalSetUp", typeof(double));


            //Get All Losses and Insert into DataTable
            DataTable LossCodesData = new DataTable();
            using (MsqlConnection mcLossCodes = new MsqlConnection())
            {
                mcLossCodes.open();
                string startDateTime1 = frmDate.ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0);
                //string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where MessageType != 'BREAKDOWN' and ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "') end) ) and LossCodeID NOT IN (  "
                //            + "SELECT DISTINCT LossCodeID FROM (  "
                //            + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where MessageType != 'BREAKDOWN' and LossCodesLevel1ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) end) ) "
                //            + "UNION  "
                //            + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where MessageType != 'BREAKDOWN' and LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                //            + "( case when (IsDeleted = 1) then ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) end) )  "
                //            + ") AS derived ) order by LossCodesLevel1ID;";


                string query = @"select LossCodeID,LossCode from i_facility_tal.dbo.tbllossescodes  where MessageType != 'BREAKDOWN' and ((CreatedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or "
                            + "(  (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "') ) ) and LossCodeID NOT IN (  "
                            + "SELECT DISTINCT LossCodeID FROM (  "
                            + "SELECT DISTINCT LossCodesLevel1ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where MessageType != 'BREAKDOWN' and LossCodesLevel1ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                            + "(  (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) ) ) "
                            + "UNION  "
                            + "SELECT DISTINCT LossCodesLevel2ID AS LossCodeID FROM i_facility_tal.dbo.tbllossescodes where MessageType != 'BREAKDOWN' and LossCodesLevel2ID is not null  and ((CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or  "
                            + "(  (IsDeleted = 1) and ( CreatedOn <=  '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ModifiedOn >= '" + startDateTime1 + "' ) ) )  "
                            + ") AS derived ) order by LossCodesLevel1ID;";

                SqlDataAdapter daLossCodesData = new SqlDataAdapter(query, mcLossCodes.msqlConnection);
                daLossCodesData.Fill(LossCodesData);
                mcLossCodes.close();
            }
            int LossesStartsATCol = 27;
            var LossesList = new List<KeyValuePair<int, string>>();

            #region LossCodes Into LossList
            for (int i = 0; i < LossCodesData.Rows.Count; i++)
            {
                int losscode = Convert.ToInt32(LossCodesData.Rows[i][0]);
                string losscodeName = Convert.ToString(LossCodesData.Rows[i][1]);

                var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == losscode).FirstOrDefault();
                int level = lossdata.LossCodesLevel;
                if (level == 3)
                {
                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                    int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                    var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                    losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                }

                else if (level == 2)
                {
                    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                    losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                }
                else if (level == 1)
                {
                    if (losscode == 999)
                    {
                        losscodeName = "NoCode Entered";
                    }
                    else if (losscode == 9999)
                    {
                        losscodeName = "UnIdentified BreakDown";
                    }
                    else
                    {
                        losscodeName = lossdata.LossCode;
                    }
                }
                //losscodeName = LossHierarchy3rdLevel(losscode);
                DTConsolidatedLosses.Columns.Add(losscodeName, typeof(double));
                DTConsolidatedLosses.Columns[losscodeName].DefaultValue = "0";

                //Code to write LossesNames to Excel.
                string columnAlphabet = ExcelColumnFromNumber(LossesStartsATCol);

                worksheet.Cells[columnAlphabet + 4].Value = losscodeName;
                //worksheet.Cells[columnAlphabet + 5].Value = "AF";

                LossesStartsATCol++;
                //Add the LossesToList
                LossesList.Add(new KeyValuePair<int, string>(losscode, losscodeName));
            }
            #endregion

            #region Push Headers that r supposed to be after Losses

            int ColIndex = LossCodesData.Rows.Count + 26 + 1; //+1 For Gap (& Testing) //26 is Previous DATA(Plant,Shop......)
            string ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "Rejected Qty";
            ColIndex = LossCodesData.Rows.Count + 26 + 2;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "Rejected Reason";
            ColIndex = LossCodesData.Rows.Count + 26 + 3;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "Operator Name";
            ColIndex = LossCodesData.Rows.Count + 26 + 4;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "Type";
            //To skip a Column Just Increment the ColIndex extra +1
            ColIndex = LossCodesData.Rows.Count + 26 + 6;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "PartNo & OpNo";
            ColIndex = LossCodesData.Rows.Count + 26 + 7;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "NC Cutting Time Per Part";
            ColIndex = LossCodesData.Rows.Count + 26 + 8;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "Total NC Cutting Time";
            ColIndex = LossCodesData.Rows.Count + 26 + 9;
            ColAfterLosses = ExcelColumnFromNumber(ColIndex);
            worksheet.Cells[ColAfterLosses + "4"].Value = "%";

            #endregion

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5; // Gap to Insert OverAll data. DataStartRow + MachinesCount + 2(1 for HighestLevel & another for Gap).
            int Sno = 1;
            string finalLossCol = null;

            for (int i = 0; i < TotalDay + 1; i++)
            {
                int StartingRowForToday = Row;
                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");

                int NumMacsToExcel = 0;
                for (int n = 0; n < machin.Rows.Count; n++)
                {
                    NumMacsToExcel++;
                    double CummulativeOfAllLosses = 0;
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        StartingRowForToday = Row;
                    }

                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);
                    List<int> MacList = new List<int>();
                    int IsNormalWC = 0;
                    IsNormalWC = Convert.ToInt32(db.tblmachinedetails.Where(m => m.MachineID == MachineID).Select(m => m.IsNormalWC).FirstOrDefault());
                    if (IsNormalWC == 1)
                    {
                        var SubWCData = db.tblmachinedetails.Where(m => m.ManualWCID == MachineID && m.IsDeleted == 0).ToList();
                        foreach (var subMacs in SubWCData)
                        {
                            MacList.Add(Convert.ToInt32(subMacs.MachineID));
                        }
                    }
                    else
                    {
                        MacList.Add(MachineID);
                    }

                    //Added this machineDetails into Datatable
                    string WCInvNoString = HierarchyData[3];
                    //string correctedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
                    string correctedDate = StartDate;

                    #region Get HMI DATA and Push For this machine and Date.
                    //Get general hmi data, later get prodfai && operator Combination based data inside loop

                    //var HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();

                    //var HMIDATA = (from r in db.tblworeports
                    //               where MacList.Contains((int)r.MachineID) && r.CorrectedDate == correctedDate
                    //               select r).ToList().GroupBy(a => a.HMIID).ToList();

                    var HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID))
                        .Where(m => m.CorrectedDate == correctedDate).OrderByDescending(m => m.HMIID).ToList();

                    if ((!string.IsNullOrEmpty(Operator.Trim())) && ProdFAI == "OverAll")
                    {
                        //HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.OperatorName == Operator).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                        HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID) && m.CorrectedDate == correctedDate && m.OperatorName == Operator)
                        .OrderByDescending(m => m.HMIID).ToList();
                    }
                    else if ((!string.IsNullOrEmpty(Operator.Trim())) && ProdFAI != "OverAll")
                    {
                        //HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.OperatorName == Operator && m.Type == ProdFAI).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                        if (ProdFAI != "Others")
                        {
                            HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID) && m.CorrectedDate == correctedDate && m.OperatorName == Operator && m.Type == ProdFAI)
                           .OrderByDescending(m => m.HMIID).ToList();
                        }
                        else
                        {
                            HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID) && m.CorrectedDate == correctedDate && m.OperatorName == Operator && m.Type != "Prod" && m.Type != "FAI")
                           .OrderByDescending(m => m.HMIID).ToList();
                        }
                    }
                    else if ((string.IsNullOrEmpty(Operator.Trim())) && ProdFAI != "OverAll")
                    {
                        //HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.Type == ProdFAI).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                        if (ProdFAI != "Others")
                        {
                            HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID) && m.CorrectedDate == correctedDate && m.Type == ProdFAI)
                         .OrderByDescending(m => m.HMIID).ToList();
                            //.OrderByDescending(m => m.HMIID).Take(1).ToList(); //Wrong on 2017-03-24
                        }
                        else
                        {
                            HMIDATA = db.tblworeports.Where(m => MacList.Contains((int)m.MachineID) && m.CorrectedDate == correctedDate && m.Type != "Prod" && m.Type != "FAI")
                        .OrderByDescending(m => m.HMIID).ToList();
                        }
                    }

                    //if ((!string.IsNullOrEmpty(Operator.Trim())) && ProdFAI == "OverAll")
                    //{
                    //    HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.OperatorName == Operator).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                    //}
                    //else if ((!string.IsNullOrEmpty(Operator.Trim())) && ProdFAI != "OverAll")
                    //{
                    //    HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.OperatorName == Operator && m.Type == ProdFAI).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                    //}
                    //else if ((string.IsNullOrEmpty(Operator.Trim())) && ProdFAI != "OverAll")
                    //{
                    //    HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.Type == ProdFAI).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                    //}


                    //if (Operator == null & ProdFAI == "OverAll") . This purpose is served by 1st time HMIDATA is assigned.
                    if (HMIDATA != null)
                    {
                        //Now Loop through and Push Data into Excel
                        foreach (var row in HMIDATA)
                        {

                            //DateTime woStartTime = Convert.ToDateTime(row.Date);
                            //DateTime woEndTime = Convert.ToDateTime(row.Time);

                            if (Convert.ToInt32(row.IsMultiWO) == 1) //Its a MultiWorkOrder
                            {
                                #region
                                int HmiID = Convert.ToInt32(row.HMIID);
                                var MulitWOData = db.tblworeports.Where(m => m.HMIID == HmiID).ToList();
                                worksheet.Cells["B" + Row + ":B" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["C" + Row + ":C" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["D" + Row + ":D" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["E" + Row + ":E" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["F" + Row + ":F" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["G" + Row + ":G" + (Row + MulitWOData.Count)].Merge = true;
                                worksheet.Cells["H" + Row + ":H" + (Row + MulitWOData.Count)].Merge = true;

                                worksheet.Cells["B" + Row].Value = Sno++;
                                worksheet.Cells["C" + Row].Value = HierarchyData[0];
                                worksheet.Cells["D" + Row].Value = HierarchyData[1];
                                worksheet.Cells["E" + Row].Value = HierarchyData[2];
                                worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                                worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV
                                worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                                worksheet.Cells["I" + Row].Value = row.Shift;
                                worksheet.Cells["J" + Row].Value = row.PartNo + "-" + MulitWOData.Count;
                                worksheet.Cells["K" + Row].Value = row.WorkOrderNo + "-" + MulitWOData.Count;
                                worksheet.Cells["L" + Row].Value = row.OpNo + "-" + MulitWOData.Count;
                                worksheet.Cells["M" + Row].Value = MulitWOData.Sum(m => m.TargetQty);
                                double delQty = 0;
                                //double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                                worksheet.Cells["N" + Row].Value = MulitWOData.Sum(m => m.DeliveredQty);
                                if (row.IsPF == 0)
                                {
                                    worksheet.Cells["O" + Row].Value = "Yes";
                                    //worksheet.Cells["P" + Row].Value = "";
                                }
                                else
                                {
                                    //worksheet.Cells["O" + Row].Value = "";
                                    worksheet.Cells["P" + Row].Value = "Yes";
                                }
                                worksheet.Cells["Q" + Row].Value = "";//isHold
                                worksheet.Cells["R" + Row].Value = "";//Hold Reason

                                double SettingTime = Convert.ToDouble(row.SettingTime);
                                worksheet.Cells["S" + Row].Value = Math.Round(SettingTime, 1);

                                double Green = Convert.ToDouble(row.CuttingTime);
                                worksheet.Cells["T" + Row].Value = Green;

                                //double Changeover = GetChangeoverTimeForWO(correctedDate, MachineID, woStartTime, woEndTime);
                                //worksheet.Cells["S" + Row].Value = Math.Round(Changeover / 60, 1);

                                double idleTime = Convert.ToDouble(row.Idle - row.SettingTime);
                                worksheet.Cells["V" + Row].Value = Math.Round(idleTime, 2);

                                worksheet.Cells["W" + Row].Formula = "=SUM(S" + Row + ",T" + Row + ",U" + Row + ",V" + Row + ")";

                                worksheet.Cells["X" + Row].Value = "";//Empty finalLossCol
                                int column = 26 + LossCodesData.Rows.Count; // StartCol in Excel + TotalLosses - 1
                                finalLossCol = ExcelColumnFromNumber(column);

                                worksheet.Cells["Y" + Row].Formula = "=SUM(Z" + Row + ":" + finalLossCol + Row + ")";

                                double BreakdownDuration = Convert.ToDouble(row.Breakdown);
                                worksheet.Cells["Z" + Row].Value = Math.Round(BreakdownDuration, 1);

                                //Push Data that is supposed to be after Losses.
                                ColIndex = LossCodesData.Rows.Count + 26 + 1; //+1 For Gap (& Testing) //26 is Previous DATA(Plant,Shop......)
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                worksheet.Cells[ColAfterLosses + Row].Value = 0;
                                ColIndex = LossCodesData.Rows.Count + 26 + 2;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                worksheet.Cells[ColAfterLosses + Row].Value = "0";
                                ColIndex = LossCodesData.Rows.Count + 26 + 3;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                worksheet.Cells[ColAfterLosses + Row].Value = row.OperatorName;
                                ColIndex = LossCodesData.Rows.Count + 26 + 4;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                worksheet.Cells[ColAfterLosses + Row].Value = row.Type;
                                //To skip a Column Just Increment the ColIndex extra +1
                                ColIndex = LossCodesData.Rows.Count + 26 + 6;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                worksheet.Cells[ColAfterLosses + Row].Value = row.PartNo + "#" + row.OpNo;
                                ColIndex = LossCodesData.Rows.Count + 26 + 7;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                string ColForNCCuttingTime = ColAfterLosses;
                                string partNo = row.PartNo;
                                string opNo = row.OpNo;
                                //double stdCuttingTime = 0;
                                //string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == partNo && m.OpNo == opNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                                //double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                                //worksheet.Cells[ColAfterLosses + Row].Value = stdCuttingTime;
                                double stdCuttingTime = 0;
                                double ProdOfstdCuttingTimeDelivQty = 0;
                                //int HmiID = row.HMIID;
                                int PF = 0, JF = 0;
                                //var MulitWOData = db.tbl_multiwoselection.Where(m => m.HMIID == HmiID);
                                foreach (var MulitWOrow in MulitWOData)
                                {
                                    double stdCuttingTimeLocal = 0;
                                    int DelivQty = 0;
                                    //string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == MulitWOrow.PartNo && m.OpNo == MulitWOrow.OperationNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                                    //string DelivQtyString = Convert.ToString(MulitWOrow.DeliveredQty);
                                    //int.TryParse(DelivQtyString, out DelivQty);
                                    //double.TryParse(stdCuttingTimeString, out stdCuttingTimeLocal);

                                    stdCuttingTimeLocal = Convert.ToDouble(MulitWOrow.CuttingTime);
                                    DelivQty = Convert.ToInt32(MulitWOrow.DeliveredQty);
                                    stdCuttingTime += stdCuttingTimeLocal;
                                    ProdOfstdCuttingTimeDelivQty += DelivQty * stdCuttingTimeLocal;

                                    if (MulitWOrow.IsPF == 0)
                                    {
                                        PF += 1;
                                    }
                                    else
                                    {
                                        JF += 1;
                                    }
                                }
                                worksheet.Cells[ColAfterLosses + Row].Value = stdCuttingTime;
                                ColIndex = LossCodesData.Rows.Count + 26 + 8;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                //worksheet.Cells[ColAfterLosses + Row].Formula = "=PRODUCT( N" + Row + "," + ColForNCCuttingTime + Row + ")";
                                //double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                                worksheet.Cells[ColAfterLosses + Row].Value = ProdOfstdCuttingTimeDelivQty;
                                string ColForTotalNCCuttingTime = ColAfterLosses;
                                ColIndex = LossCodesData.Rows.Count + 26 + 9;
                                ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                //worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")";
                                //worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ R" + Row + ")";
                                // var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")");
                                var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ R" + Row + ")");
                                double Percentage = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);
                                if (Percentage < 100)
                                {
                                    worksheet.Cells[ColAfterLosses + Row].Value = Percentage;
                                }
                                else
                                {
                                    worksheet.Cells[ColAfterLosses + Row].Value = 100;
                                }

                                // Push Loss Value into  DataTable & Excel
                                DataRow dr = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoString && r.Field<string>("CorrectedDate") == dateforMachine);
                                if (dr != null)
                                {
                                    //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                    //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                    DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                    HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, PF + JF, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                }
                                else
                                {
                                    //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                    //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                    DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                    HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, PF + JF, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                }

                                #region Capture and Push Losses

                                //now push 0 for every other loss into excel
                                worksheet.Cells["AA" + Row + ":" + finalLossCol + Row].Value = Convert.ToDouble(0.0);

                                //to Capture and Push , Losses that occured.
                                //List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSecondsForWO(MachineID, correctedDate, woStartTime, woEndTime);
                                //DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                                //if (dr1 != null)
                                //{
                                //    foreach (var loss in LossesdurationList)
                                //    {

                                //        int LossID = loss.Key;
                                //        double Duration = loss.Value;
                                //        var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossID).FirstOrDefault();
                                //        int level = lossdata.LossCodesLevel;
                                //        string losscodeName = null;

                                #region To Get LossCode Hierarchy
                                //        if (level == 3)
                                //        {
                                //            int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                //            int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                                //            var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                                //            var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                                //            losscodeName = lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
                                //        }
                                //        else if (level == 2)
                                //        {
                                //            int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                                //            var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                                //            losscodeName = lossdata1.LossCode + ":" + lossdata.LossCode;
                                //        }
                                //        else if (level == 1)
                                //        {
                                //            if (LossID == 999)
                                //            {
                                //                losscodeName = "NoCode Entered";
                                //            }
                                //            else
                                //            {
                                //                losscodeName = lossdata.LossCode;
                                //            }
                                //        }
                                #endregion

                                //        int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                //        string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 9);
                                //        double DurInMinutes = Convert.ToDouble(Math.Round((Duration / (60)), 1)); //To Minutes:: 1 Decimal Place
                                //        worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInMinutes;
                                //        if (DurInMinutes > 0)
                                //        {
                                //        }
                                //        dr1[losscodeName] = DurInMinutes;
                                //        CummulativeOfAllLosses += DurInMinutes;
                                //    }
                                //}

                                //to Capture and Push , Losses that occured.
                                // List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSecondsForWO(MachineID, correctedDate, woStartTime, woEndTime);
                                var LossesdurationList = db.tblwolossesses.Where(m => m.HMIID == HmiID).ToList();
                                DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                                if (dr1 != null)
                                {
                                    foreach (var loss in LossesdurationList)
                                    {
                                        int LossID = Convert.ToInt32(loss.LossID);
                                        double Duration = Convert.ToDouble(loss.LossDuration);
                                        int level = Convert.ToInt32(loss.Level);
                                        string losscodeName = loss.LossName;

                                        #region To Get LossCode Hierarchy
                                        if (level == 3)
                                        {
                                            losscodeName = loss.LossCodeLevel1Name + " :: " + loss.LossCodeLevel2Name + " : " + losscodeName;
                                        }
                                        else if (level == 2)
                                        {
                                            losscodeName = loss.LossCodeLevel1Name + ":" + losscodeName;
                                        }
                                        else if (level == 1)
                                        {
                                            if (LossID == 999)
                                            {
                                                losscodeName = "NoCode Entered";
                                            }
                                            else
                                            {
                                                losscodeName = losscodeName;
                                            }
                                        }
                                        #endregion

                                        int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                        string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 9);
                                        double DurInMinutes = Convert.ToDouble(Math.Round((Duration), 1)); //To Minutes:: 1 Decimal Place
                                        worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInMinutes;
                                        if (DurInMinutes > 0)
                                        {
                                        }
                                        dr1[losscodeName] = DurInMinutes;
                                        CummulativeOfAllLosses += DurInMinutes;
                                    }
                                }

                                Row++;
                                #endregion

                                //individual WO's
                                int hmiID = Convert.ToInt32(row.HMIID);
                                var HMIDATA1 = db.tbl_multiwoselection.Where(m => m.HMIID == hmiID).ToList();
                                foreach (var row1 in HMIDATA1)
                                {
                                    #region To push to excel. Multi WO.
                                    //worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                                    //worksheet.Cells["I" + Row].Value = row.Shift;
                                    worksheet.Cells["J" + Row].Value = row1.PartNo;
                                    worksheet.Cells["K" + Row].Value = row1.WorkOrder;
                                    worksheet.Cells["L" + Row].Value = row1.OperationNo;
                                    worksheet.Cells["M" + Row].Value = row1.TargetQty;
                                    delQty = 0;
                                    double.TryParse(Convert.ToString(row1.DeliveredQty), out delQty);
                                    worksheet.Cells["N" + Row].Value = delQty;
                                    if (row1.IsCompleted == 0)
                                    {
                                        worksheet.Cells["O" + Row].Value = "Yes";
                                    }
                                    else
                                    {
                                        worksheet.Cells["P" + Row].Value = "Yes";
                                    }

                                    ColIndex = LossCodesData.Rows.Count + 26 + 3;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.OperatorName;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 4;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.Type;
                                    //To skip a Column Just Increment the ColIndex extra +1
                                    ColIndex = LossCodesData.Rows.Count + 26 + 6;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row1.PartNo + "#" + row1.OperationNo;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 7;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    ColForNCCuttingTime = ColAfterLosses;
                                    partNo = row1.PartNo;
                                    opNo = row1.OperationNo;
                                    stdCuttingTime = 0;
                                    string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == partNo && m.OpNo == opNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                                    double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                                    worksheet.Cells[ColAfterLosses + Row].Value = stdCuttingTime;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 8;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=PRODUCT( N" + Row + "," + ColForNCCuttingTime + Row + ")";
                                    //double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                                    worksheet.Cells[ColAfterLosses + Row].Value = delQty * stdCuttingTime;
                                    ColForTotalNCCuttingTime = ColAfterLosses;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 9;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")";
                                    //TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")");
                                    //worksheet.Cells[ColAfterLosses + Row].Value = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime), 1);

                                    TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ R" + Row + ")");
                                    Percentage = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);
                                    if (Percentage < 100)
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = Percentage;
                                    }
                                    else
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = 100;
                                    }
                                    Row++;

                                    #endregion
                                }
                                #endregion
                            }
                            else
                            {
                                if (row.IsNormalWC == 0) //Its a NormalWC
                                {
                                    #region To push to excel. Single WO. NormalWorkCenter
                                    worksheet.Cells["B" + Row].Value = Sno++;
                                    //worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                                    worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                                    worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                                    worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                                    //worksheet.Cells["I" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                                    worksheet.Cells["I" + Row].Value = row.Shift;
                                    worksheet.Cells["J" + Row].Value = row.PartNo;
                                    worksheet.Cells["K" + Row].Value = row.WorkOrderNo;
                                    worksheet.Cells["L" + Row].Value = row.OpNo;
                                    worksheet.Cells["M" + Row].Value = row.TargetQty;
                                    double delQty = Convert.ToDouble(row.DeliveredQty);
                                    worksheet.Cells["N" + Row].Value = delQty;
                                    int PF = 0, JF = 0;//PartialFinished
                                    if (row.IsPF == 0)
                                    {
                                        //PF = 1;
                                        worksheet.Cells["O" + Row].Value = "Yes";
                                    }
                                    else
                                    {
                                        //JF = 1;
                                        worksheet.Cells["P" + Row].Value = "Yes";
                                    }

                                    worksheet.Cells["Q" + Row].Value = ""; //IsHold
                                    worksheet.Cells["R" + Row].Value = ""; //Hold Reason
                                    double SettingTime = Convert.ToDouble(row.SettingTime);
                                    worksheet.Cells["S" + Row].Value = Math.Round(SettingTime, 1);

                                    double Green = Convert.ToDouble(row.CuttingTime);
                                    worksheet.Cells["T" + Row].Value = Green;

                                    //double Changeover = GetChangeoverTimeForWO(correctedDate, MachineID, woStartTime, woEndTime);
                                    //worksheet.Cells["S" + Row].Value = Math.Round(Changeover / 60, 1);

                                    double idleTime = Convert.ToDouble(row.Idle - row.SettingTime);
                                    worksheet.Cells["V" + Row].Value = Math.Round(idleTime, 2);

                                    worksheet.Cells["W" + Row].Formula = "=SUM(S" + Row + ",T" + Row + ",U" + Row + ",V" + Row + ",Z" + Row + ")";

                                    worksheet.Cells["X" + Row].Value = "";//Empty finalLossCol
                                    int column = 26 + LossCodesData.Rows.Count; // StartCol in Excel + TotalLosses - 1
                                    finalLossCol = ExcelColumnFromNumber(column);

                                    worksheet.Cells["Y" + Row].Formula = "=SUM(Z" + Row + ":" + finalLossCol + Row + ")";

                                    double BreakdownDuration = Convert.ToDouble(row.Breakdown);
                                    worksheet.Cells["Z" + Row].Value = Math.Round(BreakdownDuration, 1);

                                    //Push Data that is supposed to be after Losses.
                                    ColIndex = LossCodesData.Rows.Count + 26 + 1; //+1 For Gap (& Testing) //26 is Previous DATA(Plant,Shop......)
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = 0;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 2;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = "0";
                                    ColIndex = LossCodesData.Rows.Count + 26 + 3;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.OperatorName;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 4;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.Type;
                                    //To skip a Column Just Increment the ColIndex extra +1
                                    ColIndex = LossCodesData.Rows.Count + 26 + 6;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.PartNo + "#" + row.OpNo;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 7;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    string ColForNCCuttingTime = ColAfterLosses;
                                    string partNo = row.PartNo;
                                    string opNo = row.OpNo;
                                    double stdCuttingTime = 0;
                                    double ProdOfstdCuttingTimeDelivQty = 0;
                                    int totalWOs = 1;
                                    double TargetNC = 0;

                                    //string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == partNo && m.OpNo == opNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                                    //double.TryParse(stdCuttingTimeString, out stdCuttingTime);

                                    stdCuttingTime = Convert.ToDouble(row.NCCuttingTimePerPart);
                                    ProdOfstdCuttingTimeDelivQty = delQty * stdCuttingTime;

                                    if (row.IsPF == 0)
                                    {
                                        PF = 1;
                                    }
                                    else
                                    {
                                        JF = 1;
                                    }

                                    worksheet.Cells[ColAfterLosses + Row].Value = stdCuttingTime;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 8;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=PRODUCT( N" + Row + "," + ColForNCCuttingTime + Row + ")";
                                    //double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                                    worksheet.Cells[ColAfterLosses + Row].Value = ProdOfstdCuttingTimeDelivQty;
                                    string ColForTotalNCCuttingTime = ColAfterLosses;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 9;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")";
                                    //var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")");
                                    worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((T" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ T" + Row + ")";
                                    var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((T" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ T" + Row + ")");

                                    double Percentage = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);
                                    if (Percentage < 100)
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = Percentage;
                                    }
                                    else
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = 100;
                                    }

                                    //worksheet.Cells[ColAfterLosses + Row].Value = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);

                                    //Now get & put Losses
                                    // Push Loss Value into  DataTable & Excel
                                    DataRow dr = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoString && r.Field<string>("CorrectedDate") == dateforMachine);
                                    if (dr != null)
                                    {
                                        //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                        //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                        DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                        HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, totalWOs, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                    }
                                    else
                                    {
                                        //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                        //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                        DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                        HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, totalWOs, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                    }

                                    #region Capture and Push Losses

                                    //now push 0 for every other loss into excel
                                    worksheet.Cells["AA" + Row + ":" + finalLossCol + Row].Value = Convert.ToDouble(0.0);

                                    //to Capture and Push , Losses that occured.
                                    //List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSecondsForWO(MachineID, correctedDate, woStartTime, woEndTime);
                                    int HmiID = Convert.ToInt32(row.HMIID);
                                    var LossesdurationList = db.tblwolossesses.Where(m => m.HMIID == HmiID).ToList();
                                    DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                                    if (dr1 != null)
                                    {
                                        foreach (var loss in LossesdurationList)
                                        {
                                            int LossID = Convert.ToInt32(loss.LossID);
                                            double Duration = Convert.ToDouble(loss.LossDuration);
                                            int level = Convert.ToInt32(loss.Level);
                                            string losscodeName = loss.LossName;

                                            #region To Get LossCode Hierarchy
                                            if (level == 3)
                                            {
                                                losscodeName = loss.LossCodeLevel1Name + " :: " + loss.LossCodeLevel2Name + " : " + losscodeName;
                                            }
                                            else if (level == 2)
                                            {
                                                losscodeName = loss.LossCodeLevel1Name + ":" + losscodeName;
                                            }
                                            else if (level == 1)
                                            {
                                                if (LossID == 999)
                                                {
                                                    losscodeName = "NoCode Entered";
                                                }
                                                else
                                                {
                                                    losscodeName = losscodeName;
                                                }
                                            }
                                            #endregion

                                            int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                            string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 9);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                                            double DurInMinutes = Convert.ToDouble(Math.Round((Duration), 1)); //To Minutes:: 1 Decimal Place
                                            worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInMinutes;
                                            dr1[losscodeName] = DurInMinutes;
                                            CummulativeOfAllLosses += DurInMinutes;
                                        }
                                    }
                                    Row++;
                                    #endregion

                                    #endregion
                                }
                                else if (row.IsNormalWC == 1) //Its a ManualWC
                                {
                                    #region To push to excel. Single WO. NormalWorkCenter
                                    worksheet.Cells["B" + Row].Value = Sno++;
                                    //worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                                    worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV
                                    worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                                    // worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                                    worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                                    //worksheet.Cells["I" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                                    worksheet.Cells["I" + Row].Value = row.Shift;
                                    worksheet.Cells["J" + Row].Value = row.PartNo;
                                    worksheet.Cells["K" + Row].Value = row.WorkOrderNo;
                                    worksheet.Cells["L" + Row].Value = row.OpNo;
                                    worksheet.Cells["M" + Row].Value = row.TargetQty;
                                    double delQty = Convert.ToDouble(row.DeliveredQty);
                                    worksheet.Cells["N" + Row].Value = delQty;
                                    int PF = 0, JF = 0;//PartialFinished
                                    if (row.IsPF == 0)
                                    {
                                        //PF = 1;
                                        worksheet.Cells["O" + Row].Value = "Yes";
                                    }
                                    else
                                    {
                                        //JF = 1;
                                        worksheet.Cells["P" + Row].Value = "Yes";
                                    }

                                    int HoldCodeID = 0;
                                    string HoldcodeIDString = row.HoldReason;
                                    int.TryParse(HoldcodeIDString, out HoldCodeID);
                                    if (HoldCodeID != 0)
                                    {
                                        worksheet.Cells["Q" + Row].Value = "Yes"; //IsHold
                                        worksheet.Cells["R" + Row].Value = GetHoldHierarchy(HoldCodeID); //Hold Reason
                                    }
                                    double SettingTime = Convert.ToDouble(row.SettingTime);
                                    worksheet.Cells["S" + Row].Value = Math.Round(SettingTime, 1);

                                    double Green = Convert.ToDouble(row.CuttingTime);
                                    worksheet.Cells["T" + Row].Value = Green;

                                    //double Changeover = GetChangeoverTimeForWO(correctedDate, MachineID, woStartTime, woEndTime);
                                    //worksheet.Cells["S" + Row].Value = Math.Round(Changeover / 60, 1);

                                    double idleTime = Convert.ToDouble(row.Idle - row.SettingTime);
                                    worksheet.Cells["V" + Row].Value = Math.Round(idleTime, 2);

                                    worksheet.Cells["W" + Row].Formula = "=SUM(S" + Row + ",T" + Row + ",U" + Row + ",V" + Row + ",Z" + Row + ")";

                                    worksheet.Cells["X" + Row].Value = "";//Empty finalLossCol
                                    int column = 26 + LossCodesData.Rows.Count; // StartCol in Excel + TotalLosses - 1
                                    finalLossCol = ExcelColumnFromNumber(column);

                                    worksheet.Cells["Y" + Row].Formula = "=SUM(Z" + Row + ":" + finalLossCol + Row + ")";

                                    double BreakdownDuration = Convert.ToDouble(row.Breakdown);
                                    worksheet.Cells["Z" + Row].Value = Math.Round(BreakdownDuration, 1);

                                    //Push Data that is supposed to be after Losses.
                                    ColIndex = LossCodesData.Rows.Count + 26 + 1; //+1 For Gap (& Testing) //26 is Previous DATA(Plant,Shop......)
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = 0;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 2;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = "0";
                                    ColIndex = LossCodesData.Rows.Count + 26 + 3;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.OperatorName;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 4;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.Type;
                                    //To skip a Column Just Increment the ColIndex extra +1
                                    ColIndex = LossCodesData.Rows.Count + 26 + 6;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    worksheet.Cells[ColAfterLosses + Row].Value = row.PartNo + "#" + row.OpNo;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 7;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    string ColForNCCuttingTime = ColAfterLosses;
                                    string partNo = row.PartNo;
                                    string opNo = row.OpNo;
                                    double stdCuttingTime = 0;
                                    double ProdOfstdCuttingTimeDelivQty = 0;
                                    int totalWOs = 1;
                                    double TargetNC = 0;

                                    //string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == partNo && m.OpNo == opNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                                    //double.TryParse(stdCuttingTimeString, out stdCuttingTime);

                                    stdCuttingTime = Convert.ToDouble(row.NCCuttingTimePerPart);
                                    ProdOfstdCuttingTimeDelivQty = delQty * stdCuttingTime;

                                    if (row.IsPF == 0)
                                    {
                                        PF = 1;
                                    }
                                    else
                                    {
                                        JF = 1;
                                    }

                                    worksheet.Cells[ColAfterLosses + Row].Value = stdCuttingTime;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 8;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=PRODUCT( N" + Row + "," + ColForNCCuttingTime + Row + ")";
                                    //double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                                    worksheet.Cells[ColAfterLosses + Row].Value = ProdOfstdCuttingTimeDelivQty;
                                    string ColForTotalNCCuttingTime = ColAfterLosses;
                                    ColIndex = LossCodesData.Rows.Count + 26 + 9;
                                    ColAfterLosses = ExcelColumnFromNumber(ColIndex);
                                    //worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")";
                                    //var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((R" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0, R" + Row + "/" + ColForTotalNCCuttingTime + Row + ")");
                                    worksheet.Cells[ColAfterLosses + Row].Formula = "=IF((OR((T" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ T" + Row + ")";
                                    var TotalNCCutTimeDIVCuttingTime = worksheet.Calculate("=IF((OR((T" + Row + "<=0),(" + ColForTotalNCCuttingTime + Row + " <=0))), 0," + ColForTotalNCCuttingTime + Row + "/ T" + Row + ")");

                                    double Percentage = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);
                                    if (Percentage < 100)
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = Percentage;
                                    }
                                    else
                                    {
                                        worksheet.Cells[ColAfterLosses + Row].Value = 100;
                                    }

                                    //worksheet.Cells[ColAfterLosses + Row].Value = Math.Round(Convert.ToDouble(TotalNCCutTimeDIVCuttingTime) * 100, 0);

                                    //Now get & put Losses
                                    // Push Loss Value into  DataTable & Excel
                                    DataRow dr = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == @WCInvNoString && r.Field<string>("CorrectedDate") == dateforMachine);
                                    if (dr != null)
                                    {
                                        //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                        //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                        DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                        HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, totalWOs, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                    }
                                    else
                                    {
                                        //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                                        //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                                        DTConsolidatedLosses.Rows.Add(HierarchyData[0], HierarchyData[1], HierarchyData[2], HierarchyData[3],
                                        HierarchyData[4], dateforMachine, row.OperatorName, PF, JF, totalWOs, row.TargetQty, delQty, ProdOfstdCuttingTimeDelivQty, Green, idleTime, SettingTime);
                                    }

                                    #region Capture and Push Losses

                                    //now push 0 for every other loss into excel
                                    worksheet.Cells["AA" + Row + ":" + finalLossCol + Row].Value = Convert.ToDouble(0.0);

                                    //to Capture and Push , Losses that occured.
                                    //List<KeyValuePair<int, double>> LossesdurationList = GetAllLossesDurationSecondsForWO(MachineID, correctedDate, woStartTime, woEndTime);
                                    int HmiID = Convert.ToInt32(row.HMIID);
                                    var LossesdurationList = db.tblwolossesses.Where(m => m.HMIID == HmiID).ToList();
                                    DataRow dr1 = DTConsolidatedLosses.AsEnumerable().FirstOrDefault(r => r.Field<string>("WCInvNo") == WCInvNoString && r.Field<string>("CorrectedDate") == @dateforMachine);
                                    if (dr1 != null)
                                    {
                                        foreach (var loss in LossesdurationList)
                                        {
                                            int LossID = Convert.ToInt32(loss.LossID);
                                            double Duration = Convert.ToDouble(loss.LossDuration);
                                            int level = Convert.ToInt32(loss.Level);
                                            string losscodeName = loss.LossName;

                                            #region To Get LossCode Hierarchy
                                            if (level == 3)
                                            {
                                                losscodeName = loss.LossCodeLevel1Name + " :: " + loss.LossCodeLevel2Name + " : " + losscodeName;
                                            }
                                            else if (level == 2)
                                            {
                                                losscodeName = loss.LossCodeLevel1Name + ":" + losscodeName;
                                            }
                                            else if (level == 1)
                                            {
                                                if (LossID == 999)
                                                {
                                                    losscodeName = "NoCode Entered";
                                                }
                                                else
                                                {
                                                    losscodeName = losscodeName;
                                                }
                                            }
                                            #endregion

                                            int ColumnIndex = DTConsolidatedLosses.Columns[losscodeName].Ordinal;
                                            string ColumnForThisLoss = ExcelColumnFromNumber(ColumnIndex + 9);// 5 is the Difference between position of Excel and DataTable Structure  for Losses Inserting column.
                                            double DurInMinutes = Convert.ToDouble(Math.Round((Duration), 1)); //To Minutes:: 1 Decimal Place
                                            worksheet.Cells[ColumnForThisLoss + "" + Row].Value = DurInMinutes;
                                            dr1[losscodeName] = DurInMinutes;
                                            CummulativeOfAllLosses += DurInMinutes;
                                        }
                                    }
                                    Row++;
                                    #endregion

                                    #endregion
                                }
                            }


                        }//end of 1 HMIDATA Row
                    }
                    #endregion

                    //worksheet.Cells["W" + Row].Value = Convert.ToDouble(Math.Round((CummulativeOfAllLosses), 1));
                    // Row++;

                }//End of For Each Machine Loop

                if (StartingRowForToday != Row)
                {
                    //Stuff for entire day (of all WC's) Into DT
                    DTConsolidatedLosses.Rows.Add("Summarized", "Summarized", "Summarized", "Summarized", "Summarized", dateforMachine);

                    //Push each Date Cummulative. Loop through ExcelAddress and insert formula
                    var rangeIndividualSummarized = worksheet.Cells["AA4:" + finalLossCol + "4"];
                    foreach (var rangeBase in rangeIndividualSummarized)
                    {
                        string str = Convert.ToString(rangeBase);
                        string ExcelColAlphabet = Regex.Replace(str, "[^A-Z _]", "");
                        worksheet.Cells[ExcelColAlphabet + Row].Formula = "=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")";
                        //var a = worksheet.Cells[rangeBase.Address].Value;
                        var blah1 = worksheet.Calculate("=SUM(" + ExcelColAlphabet + StartingRowForToday + ":" + ExcelColAlphabet + "" + (Row - 1) + ")");

                        double LossVal = 0;
                        double.TryParse(Convert.ToString(blah1), out LossVal);
                        if (LossVal != 0.0)
                        {
                            string LossName = Convert.ToString(worksheet.Cells[ExcelColAlphabet + 4].Value);
                            DataRow dr = DTConsolidatedLosses.AsEnumerable().LastOrDefault(r => r.Field<string>("Plant") == "Summarized" && r.Field<string>("CorrectedDate") == dateforMachine);
                            if (dr != null)
                            {
                                dr[LossName] = LossVal;
                            }
                        }
                    }

                    //Total of Today into 
                    //Insert Cummulative for today + 9 cols extra
                    if (Row > StartingRowForToday)
                    {
                        int col = 26 + LossCodesData.Rows.Count + 9; // StartCol of Losses + AllLosses + 9 AfterLossColumns
                        finalLossCol = ExcelColumnFromNumber(col);

                        worksheet.Cells["C" + Row + ":G" + Row].Merge = true;
                        worksheet.Cells["C" + Row].Value = "Summarized For";
                        worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd"); //For Date
                        worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells["B" + Row + ":" + finalLossCol + Row].Style.Font.Bold = true;

                        worksheet.Cells["S" + Row].Formula = "=SUM(S" + StartingRowForToday + ":S" + (Row - 1) + ")";
                        worksheet.Cells["T" + Row].Formula = "=SUM(T" + StartingRowForToday + ":T" + (Row - 1) + ")";
                        worksheet.Cells["V" + Row].Formula = "=SUM(V" + StartingRowForToday + ":V" + (Row - 1) + ")";
                        worksheet.Cells["W" + Row].Formula = "=SUM(W" + StartingRowForToday + ":W" + (Row - 1) + ")";
                        worksheet.Cells["Y" + Row].Formula = "=SUM(Y" + StartingRowForToday + ":Y" + (Row - 1) + ")";
                        worksheet.Cells["Z" + Row].Formula = "=SUM(Z" + StartingRowForToday + ":Z" + (Row - 1) + ")";

                        //Cellwise Border for Today
                        worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //Excel:: Border Around Cells.
                        worksheet.Cells["B" + StartingRowForToday + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells["B" + Row + ":" + finalLossCol + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    }
                    Row++;
                }

                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                //Row++;

            } //End of All day's Loop


            //Summarized Data (You have it all in DataTable: "Plant = 'Non-Summarized'" rows)
            Row = 5;
            Sno = 1;
            for (int n = 0; n < machin.Rows.Count; n++)
            {
                int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                List<string> HierarchyData = GetHierarchyData(MachineID);

                //get distinct opratornames for this mac & Date

                //List<string> distinctOpNames = (from row in DTConsolidatedLosses.AsEnumerable()
                //                           where row.Field<string>("WCInvNo") == HierarchyData[3][3] && row.Field<string>("CorrectedDate") == UsedDateForExcel.ToString("yyyy-MM-dd")
                //                           select(row.Field<string>("OpName")).Distinct()).ToList();

                var MacInv = HierarchyData[3];
                var idColumn = "OpName";
                // var distinctOpNames1 = DTConsolidatedLosses.DefaultView.ToTable(true, new String[] { "OpName" });
                //var distinctOpNames1 = DTConsolidatedLosses.DefaultView.ToTable(true, idColumn)
                //    .Rows
                //    .Cast<DataRow>()
                //    //.Where(row.WCInvNo == HierarchyData[3])
                //    .Select(row => row[idColumn]) //row => row[MacInv] == HierarchyData[3]
                //    .Distinct()
                //    .ToList()
                //    ;

                //var result = (from row in DTConsolidatedLosses.AsEnumerable()
                //              where row.Field<string>("WCInvNo") == HierarchyData[3]
                //              select row);

                var distinctValues = DTConsolidatedLosses.AsEnumerable()
                    .Where(row => row.Field<string>("WCInvNo") == HierarchyData[3])
                        .Select(row => new
                        {
                            opname = row.Field<string>("OpName"),
                        })
                        .Distinct();


                foreach (var row in distinctValues)
                {
                    //if (row != "{}")
                    //if(( Regex.Replace(Convert.ToString(row),"[^a-zA-Z0-9]+","")).Trim().Length > 0)
                    if (true)
                    {

                        worksheetGraph.Cells["B" + Row].Value = Sno++;
                        worksheetGraph.Cells["C" + Row].Value = HierarchyData[0];
                        worksheetGraph.Cells["D" + Row].Value = HierarchyData[1];
                        worksheetGraph.Cells["E" + Row].Value = HierarchyData[2];
                        worksheetGraph.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                        worksheetGraph.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                        worksheetGraph.Cells["H" + Row].Value = frmDate.ToString("yyyy-MM-dd") + " - " + toDate.ToString("yyyy-MM-dd");

                        worksheetGraph.Cells["I" + Row].Value = row.opname;
                        worksheetGraph.Cells["J" + Row].Value = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<int>("WOPF"));
                        worksheetGraph.Cells["K" + Row].Value = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<int>("WOProcessed"));
                        worksheetGraph.Cells["L" + Row].Value = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<int>("TotalWOQty"));

                        worksheetGraph.Cells["M" + Row].Value = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<int>("TotalTarget"));
                        worksheetGraph.Cells["N" + Row].Value = DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<int>("TotalDelivered"));
                        double summarizedCuttingTime = 0;
                        double.TryParse((DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<double>("TargetNC")).ToString()), out summarizedCuttingTime);
                        worksheetGraph.Cells["O" + Row].Value = summarizedCuttingTime;

                        var ValueAdding = Math.Round((DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<double>("TotalValueAdding"))), 0);
                        worksheetGraph.Cells["P" + Row].Value = Math.Round(ValueAdding, 0);
                        worksheetGraph.Cells["Q" + Row].Value = Math.Round((DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<double>("TotalLosses")) / 60), 0);
                        worksheetGraph.Cells["R" + Row].Value = Math.Round((DTConsolidatedLosses.AsEnumerable().Where(x => x.Field<string>("WCInvNo") == @HierarchyData[3] && x.Field<string>("OpName") == row.opname.ToString()).Sum(x => x.Field<double>("TotalSetUp")) / 60), 0);

                        ////Efficiency
                        ////worksheetGraph.Cells["S" + Row].Formula = ;
                        ////1) select rows for this date machine operator 
                        ////2) multiply individual delivered , cuttingTime
                        ////3) write to Target NC column && Calculate Efficiency

                        //double summarizedCuttingTime = 0;
                        //DataRow[] ForMacOp = DTConsolidatedLosses.Select("WCInvNo = '" + @HierarchyData[3] + "'  AND OpName = '" + row.opname.ToString() + "' ");
                        //foreach (var dr in ForMacOp)
                        //{
                        //    double delQty = 0;
                        //    double.TryParse(dr["TotalDelivered"].ToString(), out delQty);
                        //    double NCCuttingTime = 0;
                        //    double.TryParse(dr["TargetNC"].ToString(), out NCCuttingTime);
                        //    summarizedCuttingTime += delQty * (NCCuttingTime/60); 
                        //}
                        //worksheetGraph.Cells["O" + Row].Value = summarizedCuttingTime;

                        double Efficiency = Math.Round((summarizedCuttingTime / ValueAdding) * 100, 0);
                        if (Efficiency > 100)
                        {
                            worksheetGraph.Cells["S" + Row].Value = 100;
                        }
                        else
                        {
                            worksheetGraph.Cells["S" + Row].Value = Efficiency;
                        }

                        Row++;
                    }
                }
            }
            if (Row != 5)
            {
                //Cellwise Border for Summarized
                worksheetGraph.Cells["B5:S" + (Row - 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheetGraph.Cells["B5:S" + (Row - 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheetGraph.Cells["B5:S" + (Row - 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheetGraph.Cells["B5:S" + (Row - 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            //Apply style to Losses header
            int colLast = 26 + LossCodesData.Rows.Count; // StartCol of Losses + AllLosses + 9 AfterLossColumns
            finalLossCol = ExcelColumnFromNumber(colLast);

            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#32CD32");//#32CD32:lightgreen //B8C9E9
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Fill.BackgroundColor.SetColor(colFromHex);

            colLast = 26 + LossCodesData.Rows.Count + 9; // StartCol of Losses + AllLosses + 9 AfterLossColumns
            finalLossCol = ExcelColumnFromNumber(colLast);
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            worksheet.Cells["X4:" + finalLossCol + "" + 4].Style.WrapText = true;
            worksheet.Row(4).Height = 70;
            worksheet.View.ShowGridLines = false;
            worksheetGraph.View.ShowGridLines = false;
            //worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            //worksheet.Column.AutoFit();
            //worksheet.Cells.AutoFitColumns();
            //worksheetGraph.Cells[worksheet.Dimension.Address].AutoFitColumns();
            //worksheetGraph.Cells["A3:R100"].Style.Font.Color.SetColor(Color.White);

            #region Save and Download
            p.Save();

            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "JobReport" + lowestLevelName.Trim()  + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "JobReport" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
            #endregion

        }

        // Part Learning  
        public async void PartLearningReportExcel(string StartDate, string EndDate, string PartsList, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;

            #region MacCount & LowestLevel
            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            #endregion

            #region Excel and Stuff

            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\PartLearning.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "PartLearning" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "PartLearning" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetGraph = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            if (worksheetGraph == null)
            {
                try{
                    worksheetGraph = p.Workbook.Worksheets.Add("Summarized", TemplateGraph);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            #endregion

            

            #region Get Machines List
            DataTable machin = new DataTable();
            DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime = frmDate.ToString("yyyy-MM-dd");
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = null;
                if (lowestLevel == "Plant")
                {
                    //query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + "  and IsNormalWC = 0  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + "  and IsNormalWC = 0  and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "Shop")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + "  and IsNormalWC = 0   and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "Cell")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + "  and IsNormalWC = 0  and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                else if (lowestLevel == "WorkCentre")
                {
                    //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                    query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + "  and IsNormalWC = 0  and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
                }
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(machin);
                mc.close();
            }
            #endregion

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5; // Gap to Insert OverAll data. DataStartRow + MachinesCount + 2(1 for HighestLevel & another for Gap).
            int Sno = 1;
            string finalLossCol = null;
            string existingPartNo = PartsList;

            DataTable DTWONoCuttingTime = new DataTable();
            DTWONoCuttingTime.Columns.Add("CorrectedDate", typeof(string));
            DTWONoCuttingTime.Columns.Add("WONo", typeof(string));
            DTWONoCuttingTime.Columns.Add("CuttingTime", typeof(double));

            DataTable DTAll3 = new DataTable();
            DTAll3.Columns.Add("CorrectedDate", typeof(string));
            DTAll3.Columns.Add("WONo", typeof(string));
            DTAll3.Columns.Add("PartNo", typeof(string));
            DTAll3.Columns.Add("OpNo", typeof(int));
            DTAll3.Columns.Add("CuttingTime", typeof(double));

            //DataTable for Consolidated Data 
            DataTable DTConsolidatedPL = new DataTable();
            DTConsolidatedPL.Columns.Add("CorrectedDate", typeof(string));
            DTConsolidatedPL.Columns.Add("Target", typeof(int));
            DTConsolidatedPL.Columns.Add("Delivered", typeof(int));
            DTConsolidatedPL.Columns.Add("Rejected", typeof(int));
            DTConsolidatedPL.Columns.Add("ActualCutTime", typeof(double));
            DTConsolidatedPL.Columns.Add("NCCutTime", typeof(double));
            DTConsolidatedPL.Columns.Add("SetupTime", typeof(double));
            DTConsolidatedPL.Columns.Add("SelfInspection", typeof(double));
            DTConsolidatedPL.Columns.Add("WorkOrderNo", typeof(string));

            for (int i = 0; i < TotalDay + 1; i++)
            {
                int StartingRowForToday = Row;
                //string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");
                string correctedDate = UsedDateForExcel.ToString("yyyy-MM-dd");
                int NumMacsToExcel = 0;
                #region 2017-03-24 Logic Change , we will get all details from tblworeport table.
                //for (int n = 0; n < machin.Rows.Count; n++)
                //{
                //    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                //    List<string> HierarchyData = GetHierarchyData(MachineID);
                //    string WCInvNoString = HierarchyData[3];
                //    NumMacsToExcel++;
                //    double CummulativeOfAllLosses = 0;

                //    //Now Loop through and Push Data into Excel
                //    using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
                //    {
                //        var HMIDATA = dbhmi.tblhmiscreens.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate && m.isWorkInProgress == 1).ToList();
                //        if (HMIDATA != null)
                //        {
                //            //Now Loop through and Push Data into Excel
                //            foreach (var row in HMIDATA)
                //            {
                //                if (row.IsMultiWO == 0)
                //                {
                //                    #region To push to excel. Single WO.

                //                    //check if we want to show this partno in Excel
                //                    string partnumber = row.PartNo.ToString();
                //                    if (string.Compare(existingPartNo, partnumber, true) == 0)
                //                    {
                //                        if (n == 0 && i != 0)
                //                        {
                //                            Row++;
                //                            StartingRowForToday = Row;
                //                        }
                //                        int HmiID = row.HMIID;
                //                        DateTime woStartTime = Convert.ToDateTime(row.Date);
                //                        DateTime woEndTime = Convert.ToDateTime(row.Time);

                //                        //For Testing
                //                        //if (Sno == 5)
                //                        //{
                //                        //}
                //                        worksheet.Cells["B" + Row].Value = Sno++;
                //                        worksheet.Cells["C" + Row].Value = HierarchyData[0];//Plant
                //                        worksheet.Cells["D" + Row].Value = HierarchyData[1];//Shop
                //                        worksheet.Cells["E" + Row].Value = HierarchyData[2];//Cell
                //                        worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                //                        worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                //                        worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                //                        string WorkOrderNo = row.Work_Order_No;
                //                        worksheet.Cells["I" + Row].Value = row.Work_Order_No;
                //                        worksheet.Cells["J" + Row].Value = row.PartNo;
                //                        string PartNoString = Convert.ToString(row.PartNo);
                //                        worksheet.Cells["K" + Row].Value = row.OperationNo;
                //                        int OperationNoInt = Convert.ToInt32(row.OperationNo);

                //                        worksheet.Cells["L" + Row].Value = row.Project;
                //                        worksheet.Cells["M" + Row].Value = row.Target_Qty;
                //                        double delQty = 0;
                //                        double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                //                        double processedQty = 0;
                //                        double.TryParse(Convert.ToString(row.ProcessQty), out processedQty);
                //                        worksheet.Cells["N" + Row].Value = delQty + processedQty;

                //                        string rejectedString = Convert.ToString(row.Rej_Qty);
                //                        int rejectedQty = 0;
                //                        int.TryParse(rejectedString, out rejectedQty);

                //                        worksheet.Cells["O" + Row].Value = rejectedQty; //Rejected Qty

                //                        double Green = GetGreen(correctedDate, woStartTime, woEndTime, MachineID);
                //                        double SettingTime = GetSettingTimeForWO(correctedDate, MachineID, woStartTime, woEndTime);
                //                        double SelfInsepectionTime = GetSelfInsepectionForWO(correctedDate, MachineID, woStartTime, woEndTime);

                //                        double CuttingTime = 0, SettingTimeSum = 0, SelfInsepectionTimeSum = 0;
                //                        CuttingTime = Green;
                //                        SettingTimeSum = SettingTime;
                //                        SelfInsepectionTimeSum = SelfInsepectionTime;

                //                        DataRow drGraphOuter = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                        if (drGraphOuter == null)
                //                        {
                //                            DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, Green);
                //                        }
                //                        else
                //                        {
                //                            drGraphOuter["CuttingTime"] = Convert.ToDouble(drGraphOuter["CuttingTime"]) + Green;
                //                        }

                //                        //2017-03-24
                //                        DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, Green);

                //                        //For Overall graph
                //                        //if its PF'ed previously in single WO
                //                        var HMIDATAPF = dbhmi.tblhmiscreens.Where(m => m.PartNo == row.PartNo && m.Work_Order_No == row.Work_Order_No && m.OperationNo == row.OperationNo && m.isWorkInProgress == 0).ToList();
                //                        if (HMIDATAPF != null)
                //                        {
                //                            double CuttingTimeLocal = 0, SettingTimeLocal = 0, SelfInsepectionTimeLocal = 0;
                //                            foreach (var rowPF in HMIDATAPF)
                //                            {
                //                                double GreenInner = GetGreen(correctedDate, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time), MachineID);
                //                                CuttingTimeLocal += GreenInner;

                //                                double SettingTimeInner = GetSettingTimeForWO(correctedDate, MachineID, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time));
                //                                SettingTimeLocal += SettingTimeInner;

                //                                double SelfInsepectionTimeInner = GetSelfInsepectionForWO(correctedDate, MachineID, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time));
                //                                SelfInsepectionTimeLocal += SelfInsepectionTimeInner;
                //                            }

                //                            CuttingTime += CuttingTimeLocal;
                //                            SettingTimeSum += SettingTimeLocal;
                //                            SelfInsepectionTimeSum += SelfInsepectionTimeLocal;

                //                            DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                            if (drGraph == null)
                //                            {
                //                                DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, CuttingTimeLocal);
                //                            }
                //                            else
                //                            {
                //                                drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTimeLocal;
                //                            }
                //                            //2017-03-24
                //                            DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, CuttingTimeLocal);
                //                        }

                //                        //If its PF'ed in MultiWO
                //                        //For Overall graph
                //                        var MulitWOData = dbhmi.tbl_multiwoselection.Where(m => m.PartNo == row.PartNo && m.WorkOrder == row.Work_Order_No && m.OperationNo == row.OperationNo && m.IsCompleted == 0).ToList();
                //                        foreach (var MulitWOrow in MulitWOData)
                //                        {
                //                            if (string.Compare(existingPartNo, MulitWOrow.PartNo, true) == 0)
                //                            {
                //                                worksheet.Cells["I" + Row].Value = MulitWOrow.WorkOrder;
                //                                WorkOrderNo = MulitWOrow.WorkOrder;
                //                                worksheet.Cells["J" + Row].Value = MulitWOrow.PartNo;
                //                                string PartNoString1 = MulitWOrow.PartNo;
                //                                worksheet.Cells["K" + Row].Value = MulitWOrow.OperationNo;
                //                                int OperationNoint = Convert.ToInt32(MulitWOrow.OperationNo);

                //                                #region ////For Overall graph if it was PF 'ed in HMIScreen. Its is covered in previous block of code
                //                                //var HMIDATAPFsingle = db.tblhmiscreens.Where(m => m.Work_Order_No == MulitWOrow.WorkOrder && m.PartNo == MulitWOrow.PartNo && m.OperationNo == MulitWOrow.OperationNo && m.isWorkInProgress == 0).OrderByDescending(m=>m.CorrectedDate).ToList();
                //                                //if (HMIDATAPFsingle != null)
                //                                //{
                //                                //    double CuttingTimeLocal = 0;
                //                                //    foreach (var rowPF in HMIDATAPFsingle)
                //                                //    {
                //                                //        double GreenInner = GetGreen(correctedDate, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time), MachineID);
                //                                //        CuttingTime += GreenInner;
                //                                //    }

                //                                //    DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                                //    if (drGraph == null)
                //                                //    {
                //                                //        DTWONoCuttingTime.Rows.Add(correctedDate, WorkOrderNo, CuttingTimeLocal);
                //                                //    }
                //                                //    else
                //                                //    {
                //                                //        drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTimeLocal;
                //                                //    }

                //                                //    DTAll3.Rows.Add(correctedDate, WorkOrderNo, PartNoString1, OperationNoint, CuttingTimeLocal);
                //                                //}


                //                                //For Overall graph if it was PF'ed in MultiWO .
                //                                //var HMIDATAPFMulti = db.tbl_multiwoselection.Where(m => m.WorkOrder == MulitWOrow.WorkOrder && m.PartNo == MulitWOrow.PartNo && m.OperationNo == MulitWOrow.OperationNo).ToList();
                //                                //if (HMIDATAPFMulti != null)
                //                                //{

                //                                //foreach (var rowPF in HMIDATAPFMulti)
                //                                //{
                //                                //    int hmiid = Convert.ToInt32(rowPF.HMIID);
                //                                //    var HMIDATAPFInner = db.tblhmiscreens.Where(m => m.HMIID == hmiid).FirstOrDefault();
                //                                //    if (HMIDATAPFInner != null)
                //                                //    {
                //                                //        CuttingTimeLocal += GetGreen(correctedDate, Convert.ToDateTime(HMIDATAPFInner.Date), Convert.ToDateTime(HMIDATAPFInner.Time), MachineID);
                //                                //    }
                //                                //}
                //                                #endregion

                //                                //now get the green time for this hmiid
                //                                int InnerHMIID = (int)MulitWOrow.HMIID;
                //                                var InnerHMIIDData = db.tblhmiscreens.Where(m => m.HMIID == InnerHMIID).FirstOrDefault();

                //                                double CuttingTimeLocal = GetGreen(correctedDate, Convert.ToDateTime(InnerHMIIDData.Date), Convert.ToDateTime(InnerHMIIDData.Time), MachineID);
                //                                CuttingTime += CuttingTimeLocal;

                //                                DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo);
                //                                if (drGraph == null)
                //                                {
                //                                    DTWONoCuttingTime.Rows.Add(correctedDate, MulitWOrow.WorkOrder, CuttingTimeLocal);
                //                                }
                //                                else
                //                                {
                //                                    drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTimeLocal;
                //                                }
                //                                //2017-03-24
                //                                DTAll3.Rows.Add(correctedDate, MulitWOrow.WorkOrder, MulitWOrow.PartNo, MulitWOrow.OperationNo, CuttingTimeLocal);
                //                                //}
                //                            }
                //                        }


                //                        // DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, CuttingTime);
                //                        worksheet.Cells["R" + Row].Value = Math.Round(CuttingTime, 1);

                //                        //worksheet.Cells["S" + Row].Formula = "=R"+ Row + "/N"+Row ;
                //                        double AvgCuttingTime = Math.Round(CuttingTime / (delQty + processedQty), 2);
                //                        worksheet.Cells["S" + Row].Value = AvgCuttingTime;

                //                        DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, AvgCuttingTime);

                //                        worksheet.Cells["P" + Row].Value = Math.Round(SettingTimeSum / 60, 1);
                //                        worksheet.Cells["Q" + Row].Value = Math.Round(SelfInsepectionTimeSum / 60, 1);

                //                        double stdCuttingTime = 0;
                //                        string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == row.PartNo && m.OpNo == row.OperationNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                //                        double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                //                        string stdCuttingTimeUnitString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == row.PartNo && m.OpNo == row.OperationNo).Select(m => m.StdCuttingTimeUnit).FirstOrDefault());
                //                        if (stdCuttingTimeUnitString == "Hrs")
                //                        {
                //                            stdCuttingTime = stdCuttingTime * 60;
                //                        }
                //                        else if (stdCuttingTimeUnitString == "Sec")
                //                        {
                //                            stdCuttingTime = stdCuttingTime / 60;
                //                        } //else its in minutes.

                //                        worksheet.Cells["T" + Row].Value = Math.Round(stdCuttingTime, 2); //To Hours

                //                        Row++; //2017-03-21

                //                        // Push Loss Value into  DataTable & Excel based on WO and Date
                //                        DataRow dr = DTConsolidatedPL.AsEnumerable().FirstOrDefault(r => r.Field<string>("CorrectedDate") == correctedDate && r.Field<string>("WorkOrderNo") == WorkOrderNo);
                //                        if (dr == null)
                //                        {
                //                            //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                //                            //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                //                            DTConsolidatedPL.Rows.Add(correctedDate, row.Target_Qty, delQty + processedQty, 0, CuttingTime, stdCuttingTime, SettingTimeSum, SelfInsepectionTimeSum, WorkOrderNo);
                //                        }
                //                        else
                //                        {
                //                            int targ1 = Convert.ToInt32(dr["Target"]);
                //                            int targ2 = Convert.ToInt32(row.Target_Qty);
                //                            int NewTargetQty = targ1 + targ2;
                //                            int NewdelivQty = Convert.ToInt32(dr["Delivered"]) + Convert.ToInt32(delQty + processedQty);
                //                            Double NewActualCutTime = Convert.ToDouble(dr["ActualCutTime"]) + Convert.ToDouble(CuttingTime);
                //                            double NewNCCutTime = Convert.ToDouble(dr["NCCutTime"]) + Convert.ToDouble(stdCuttingTime);
                //                            double NewSetupTime = Convert.ToDouble(dr["SetupTime"]) + Convert.ToDouble(SettingTimeSum);
                //                            double NewSelfInspectionint = Convert.ToDouble(dr["SelfInspection"]) + Convert.ToDouble(SelfInsepectionTimeSum);

                //                            dr["Target"] = NewTargetQty;
                //                            dr["Delivered"] = NewdelivQty;
                //                            dr["ActualCutTime"] = NewActualCutTime;
                //                            dr["NCCutTime"] = NewNCCutTime;
                //                            dr["SetupTime"] = NewSetupTime;
                //                            dr["SelfInspection"] = NewSelfInspectionint;
                //                        }
                //                    }
                //                    #endregion
                //                }
                //                else
                //                {
                //                    #region Its a MultiWorkOrder

                //                    if (n == 0 && i != 0)
                //                    {
                //                        Row++;
                //                        StartingRowForToday = Row;
                //                    }

                //                    DateTime woStartTime = Convert.ToDateTime(row.Date);
                //                    DateTime woEndTime = Convert.ToDateTime(row.Time);

                //                    int HmiID = row.HMIID;
                //                    var MulitWOData = dbhmi.tbl_multiwoselection.Where(m => m.HMIID == HmiID).ToList();
                //                    List<string> partnumber = new List<string>();
                //                    foreach (var rowMulitWOData in MulitWOData)
                //                    {
                //                        partnumber.Add(rowMulitWOData.PartNo.ToString());
                //                    }

                //                    //bool blah0 = existingPartNo.Intersect(partnumber).Any();
                //                    //bool success = existingPartNo.Any(a => partnumber.Any(b => b == a));
                //                    bool containsStatus = partnumber.Contains(existingPartNo, StringComparer.CurrentCultureIgnoreCase);
                //                    if (containsStatus)
                //                    {
                //                        int MainWorkOrderRow = Row;
                //                        worksheet.Cells["B" + Row + ":B" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["C" + Row + ":C" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["D" + Row + ":D" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["E" + Row + ":E" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["F" + Row + ":F" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["G" + Row + ":G" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["H" + Row + ":H" + (Row + MulitWOData.Count)].Merge = true;
                //                        worksheet.Cells["L" + Row + ":L" + (Row + MulitWOData.Count)].Merge = true;

                //                        worksheet.Cells["B" + Row].Value = Sno++;
                //                        worksheet.Cells["C" + Row].Value = HierarchyData[0];//Plant
                //                        worksheet.Cells["D" + Row].Value = HierarchyData[1];//Shop
                //                        worksheet.Cells["E" + Row].Value = HierarchyData[2];//Cell
                //                        worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                //                        worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV
                //                        worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
                //                        worksheet.Cells["I" + Row].Value = row.Work_Order_No;
                //                        string WorkOrderNo = row.Work_Order_No;
                //                        worksheet.Cells["J" + Row].Value = row.PartNo;
                //                        string PartNoS = row.PartNo;
                //                        worksheet.Cells["K" + Row].Value = row.OperationNo;
                //                        string OperationNoS = row.OperationNo;
                //                        worksheet.Cells["L" + Row].Value = row.Project;
                //                        worksheet.Cells["M" + Row].Value = row.Target_Qty;
                //                        double delQty = 0;
                //                        double.TryParse(Convert.ToString(row.Delivered_Qty), out delQty);
                //                        double processedQty = 0;
                //                        double.TryParse(Convert.ToString(row.ProcessQty), out processedQty);
                //                        worksheet.Cells["N" + Row].Value = delQty + processedQty;
                //                        worksheet.Cells["O" + Row].Value = 0; //Rejected Qty

                //                        double Green = GetGreen(correctedDate, woStartTime, woEndTime, MachineID);
                //                        double SettingTime = GetSettingTimeForWO(correctedDate, MachineID, woStartTime, woEndTime);
                //                        double SelfInsepectionTime = GetSelfInsepectionForWO(correctedDate, MachineID, woStartTime, woEndTime);

                //                        double CuttingTimeSum = 0, SettingTimeSum = 0, SelfInsepectionTimeSum = 0;
                //                        CuttingTimeSum = Green;
                //                        SettingTimeSum = SettingTime;
                //                        SelfInsepectionTimeSum = SelfInsepectionTime;

                //                        //1) IF Same WONo, PartNo, OPNo exists which is PF'ed.
                //                        //2) for individual PartNo, OPNo singly PF'ed or with Other WO's ([MultiWO]) (Basically in MultiSelectionTable)
                //                        //Handled in 2). //3) for individual PartNo, OPNo if its PF'ed with other WO's [MultiWO]

                //                        //Entire MultiWO , previously pf'ed
                //                        using (i_facility_talEntities1 dbhmi1 = new i_facility_talEntities1())
                //                        {
                //                            var HMIDATAPF1 = dbhmi1.tblhmiscreens.Where(m => m.Work_Order_No == row.Work_Order_No && m.PartNo == row.PartNo && m.OperationNo == row.OperationNo && m.isWorkInProgress == 0).ToList();
                //                            if (HMIDATAPF1 != null)
                //                            {
                //                                double CuttingTimeLocal = 0, SettingTimeLocal = 0, SelfInsepectionTimeLocal = 0;
                //                                foreach (var rowPF in HMIDATAPF1)
                //                                {
                //                                    double GreenInner = GetGreen(correctedDate, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time), MachineID);
                //                                    CuttingTimeLocal += GreenInner;

                //                                    double SettingTimeInner = GetSettingTimeForWO(correctedDate, MachineID, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time));
                //                                    SettingTimeLocal += SettingTimeInner;

                //                                    double SelfInsepectionTimeInner = GetSelfInsepectionForWO(correctedDate, MachineID, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time));
                //                                    SelfInsepectionTimeLocal += SelfInsepectionTimeInner;

                //                                    //2017-03-24
                //                                    //DTAll3.Rows.Add(rowPF.CorrectedDate, rowPF.Work_Order_No, rowPF.PartNo, rowPF.OperationNo, AvgCuttingTime);
                //                                }

                //                                CuttingTimeSum += CuttingTimeLocal;
                //                                SettingTimeSum += SettingTimeLocal;
                //                                SelfInsepectionTimeSum += SelfInsepectionTimeLocal;

                //                                DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                                if (drGraph == null)
                //                                {
                //                                    DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, CuttingTimeLocal);
                //                                }
                //                                else
                //                                {
                //                                    drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTimeLocal;
                //                                }
                //                                //2017-03-24
                //                                DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, CuttingTimeLocal);

                //                            }
                //                        }

                //                        double stdCuttingTime = 0;
                //                        double ProdOfstdCuttingTimeDelivQty = 0;
                //                        double MR = 0;
                //                        //Each Individual WO's based on Hmiid.
                //                        foreach (var MulitWOrow in MulitWOData)
                //                        {
                //                            if (string.Compare(existingPartNo, MulitWOrow.PartNo, true) == 0)
                //                            {
                //                                double stdCuttingTimeLocal = 0;
                //                                string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == MulitWOrow.PartNo && m.OpNo == MulitWOrow.OperationNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                //                                string DelivQtyString = Convert.ToString(MulitWOrow.DeliveredQty);
                //                                int DelivQty = 0;
                //                                int.TryParse(DelivQtyString, out DelivQty);
                //                                double.TryParse(stdCuttingTimeString, out stdCuttingTimeLocal);
                //                                stdCuttingTime += stdCuttingTimeLocal;
                //                                ProdOfstdCuttingTimeDelivQty += DelivQty * stdCuttingTimeLocal;
                //                            }
                //                        }

                //                        worksheet.Cells["P" + Row].Value = Math.Round(SettingTimeSum / 60, 1);
                //                        worksheet.Cells["Q" + Row].Value = Math.Round(SelfInsepectionTimeSum / 60, 1);
                //                        worksheet.Cells["R" + Row].Value = CuttingTimeSum;

                //                        double AvgCuttingTime = (CuttingTimeSum) / (delQty + processedQty);
                //                        worksheet.Cells["S" + Row].Value = Math.Round(AvgCuttingTime, 2);

                //                        // Push Loss Value into  DataTable & Excel
                //                        DataRow dr = DTConsolidatedPL.AsEnumerable().FirstOrDefault(r => r.Field<string>("CorrectedDate") == correctedDate && r.Field<string>("WorkOrderNo") == WorkOrderNo);
                //                        if (dr == null)
                //                        {
                //                            //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                //                            //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                //                            DTConsolidatedPL.Rows.Add(correctedDate, row.Target_Qty, delQty + processedQty, 0, Green, ProdOfstdCuttingTimeDelivQty, SettingTime, SelfInsepectionTime, WorkOrderNo);
                //                        }
                //                        else
                //                        {
                //                            int NewTargetQty = Convert.ToInt32(dr["Target"]) + Convert.ToInt32(row.Target_Qty);
                //                            int NewdelivQty = Convert.ToInt32(dr["Delivered"]) + Convert.ToInt32(delQty + processedQty);
                //                            double NewActualCutTime = Convert.ToDouble(dr["ActualCutTime"]) + Convert.ToDouble(Green);
                //                            double NewNCCutTime = Convert.ToDouble(dr["NCCutTime"]) + Convert.ToDouble(ProdOfstdCuttingTimeDelivQty);
                //                            double NewSetupTime = Convert.ToDouble(dr["SetupTime"]) + Convert.ToDouble(SettingTime);
                //                            double NewSelfInspectionint = Convert.ToDouble(dr["SelfInspection"]) + Convert.ToDouble(SelfInsepectionTime);

                //                            dr["Target"] = NewTargetQty;
                //                            dr["Delivered"] = NewdelivQty;
                //                            dr["ActualCutTime"] = NewActualCutTime;
                //                            dr["NCCutTime"] = NewNCCutTime;
                //                            dr["SetupTime"] = NewSetupTime;
                //                            dr["SelfInspection"] = NewSelfInspectionint;
                //                        }

                //                        DataRow drGraphOuter = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                        if (drGraphOuter == null)
                //                        {
                //                            DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, Green);
                //                        }
                //                        else
                //                        {
                //                            drGraphOuter["CuttingTime"] = Convert.ToDouble(drGraphOuter["CuttingTime"]) + Green;
                //                        }

                //                        Row++;
                //                        foreach (var MulitWOrow in MulitWOData)
                //                        {
                //                            if (string.Compare(existingPartNo, MulitWOrow.PartNo, true) == 0)
                //                            {
                //                                worksheet.Cells["I" + Row].Value = MulitWOrow.WorkOrder;
                //                                WorkOrderNo = MulitWOrow.WorkOrder;
                //                                worksheet.Cells["J" + Row].Value = MulitWOrow.PartNo;
                //                                string PartNoString = MulitWOrow.PartNo;
                //                                worksheet.Cells["K" + Row].Value = MulitWOrow.OperationNo;
                //                                int OperationNoInt = Convert.ToInt32(MulitWOrow.OperationNo);

                //                                //For Overall graph if it was PF 'ed in HMIScreen.
                //                                var HMIDATAPF = db.tblhmiscreens.Where(m => m.Work_Order_No == MulitWOrow.WorkOrder && m.PartNo == MulitWOrow.PartNo && m.OperationNo == MulitWOrow.OperationNo && m.isWorkInProgress == 0).ToList();
                //                                if (HMIDATAPF != null)
                //                                {
                //                                    double CuttingTime = 0;
                //                                    foreach (var rowPF in HMIDATAPF)
                //                                    {
                //                                        double GreenInner = GetGreen(correctedDate, Convert.ToDateTime(rowPF.Date), Convert.ToDateTime(rowPF.Time), MachineID);
                //                                        CuttingTime += GreenInner;
                //                                    }

                //                                    DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo && r.Field<string>("CorrectedDate") == correctedDate);
                //                                    if (drGraph == null)
                //                                    {
                //                                        DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, CuttingTime);
                //                                    }
                //                                    else
                //                                    {
                //                                        drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTime;
                //                                    }

                //                                    DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, Green);
                //                                }

                //                                //For Overall graph if it was PF'ed in MultiWO .
                //                                var HMIDATAPFMulti = db.tbl_multiwoselection.Where(m => m.WorkOrder == MulitWOrow.WorkOrder && m.PartNo == MulitWOrow.PartNo && m.OperationNo == MulitWOrow.OperationNo).ToList();
                //                                if (HMIDATAPFMulti != null)
                //                                {
                //                                    double CuttingTime = 0;
                //                                    foreach (var rowPF in HMIDATAPFMulti)
                //                                    {
                //                                        int hmiid = Convert.ToInt32(rowPF.HMIID);
                //                                        var HMIDATAPFInner = db.tblhmiscreens.Where(m => m.HMIID == hmiid).FirstOrDefault();
                //                                        if (HMIDATAPFInner != null)
                //                                        {
                //                                            CuttingTime += (Convert.ToDateTime(HMIDATAPFInner.Time)).Subtract(Convert.ToDateTime(HMIDATAPFInner.Date)).TotalMinutes;
                //                                        }
                //                                    }

                //                                    DataRow drGraph = DTWONoCuttingTime.AsEnumerable().FirstOrDefault(r => r.Field<string>("WONo") == WorkOrderNo);
                //                                    if (drGraph == null)
                //                                    {
                //                                        DTWONoCuttingTime.Rows.Add(correctedDate, row.Work_Order_No, CuttingTime);
                //                                    }
                //                                    else
                //                                    {
                //                                        drGraph["CuttingTime"] = Convert.ToDouble(drGraph["CuttingTime"]) + CuttingTime;
                //                                    }

                //                                    //2017-03-24
                //                                    //DTAll3.Rows.Add(correctedDate, row.Work_Order_No, row.PartNo, row.OperationNo, Green);
                //                                }

                //                                worksheet.Cells["M" + Row].Value = MulitWOrow.TargetQty;
                //                                delQty = 0;
                //                                double.TryParse(Convert.ToString(MulitWOrow.DeliveredQty), out delQty);
                //                                worksheet.Cells["N" + Row].Value = delQty;
                //                                worksheet.Cells["O" + Row].Value = 0; //Rejected Qty

                //                                stdCuttingTime = 0;
                //                                string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == MulitWOrow.PartNo && m.OpNo == MulitWOrow.OperationNo).Select(m => m.StdCuttingTime).FirstOrDefault());
                //                                double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                //                                // worksheet.Cells["Q" + Row].Value = stdCuttingTime * delQty;
                //                                worksheet.Cells["T" + Row].Value = stdCuttingTime;
                //                                Row++;
                //                            }
                //                        }
                //                    }
                //                    #endregion
                //                }
                //            }
                //        }
                //    }
                //}//for each machine Ends here for 1-Day

                #endregion

                //1) Get distinct partno,WoNo,Opno which are JF
                //2) Get sum of green, settingTime, etc and push into excel
                DataTable PartData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = "select * from [i_facility_tal].[dbo].tblworeport where HMIID in ( SELECT Max(HMIID) FROM i_facility_tal.dbo.tblworeport where PartNo = '" + existingPartNo + "' and IsPF = 1 and CorrectedDate = '" + correctedDate + "' group by WorkOrderNo,OpNo );";
                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(PartData);
                    mc.close();
                }
                for (int j = 0; j < PartData.Rows.Count; j++)
                {
                    int MachineID = Convert.ToInt32(PartData.Rows[j][1]); //MachineID
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];//Plant
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];//Shop
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];//Cell
                    worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                    worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                    worksheet.Cells["H" + Row].Value = PartData.Rows[j][5];//CorrectedDate
                    string WorkOrderNo = Convert.ToString(PartData.Rows[j][7]);//
                    worksheet.Cells["I" + Row].Value = WorkOrderNo;
                    string PartNoString = Convert.ToString(PartData.Rows[j][6]);
                    worksheet.Cells["J" + Row].Value = PartNoString;
                    string OperationNoString = Convert.ToString(PartData.Rows[j][8]);
                    int OperationNoInt = Convert.ToInt32(OperationNoString);
                    worksheet.Cells["K" + Row].Value = OperationNoString;
                    worksheet.Cells["L" + Row].Value = Convert.ToString(PartData.Rows[j][24]); //project or program
                    worksheet.Cells["M" + Row].Value = Convert.ToInt32(PartData.Rows[j][9]); //Target

                    //Its a JF and Summation of Cutting time and Stuff so. Target = Delivered 
                    double delQty = Convert.ToInt32(PartData.Rows[j][9]);
                    string rejectedString = Convert.ToString(PartData.Rows[j][22]);
                    int rejectedQty = 0;
                    int.TryParse(rejectedString, out rejectedQty);
                    worksheet.Cells["O" + Row].Value = rejectedQty; //Rejected Qty

                    double CuttingTime = 0, SettingTime = 0, SelfInsepectionTime = 0;
                    //Get PF'ed Rows for this PartNo,WONo,OpNo Combo.
                    var PFedData = db.tblworeports.Where(m => m.PartNo == PartNoString && m.WorkOrderNo == WorkOrderNo && m.OpNo == OperationNoString).ToList();
                    foreach (var row in PFedData)
                    {
                        CuttingTime += Convert.ToDouble(row.CuttingTime);
                        SettingTime += Convert.ToDouble(row.SettingTime);
                        SelfInsepectionTime += Convert.ToDouble(row.SelfInspection);
                    }

                    worksheet.Cells["R" + Row].Value = Math.Round(CuttingTime, 1);
                    worksheet.Cells["P" + Row].Value = Math.Round(SettingTime, 1);
                    worksheet.Cells["Q" + Row].Value = Math.Round(SelfInsepectionTime, 1);
                    double AvgCuttingTime = Math.Round(CuttingTime / (delQty), 2);
                    worksheet.Cells["S" + Row].Value = AvgCuttingTime;

                    double stdCuttingTime = 0;
                    string stdCuttingTimeString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == PartNoString && m.OpNo == OperationNoString).Select(m => m.StdCuttingTime).FirstOrDefault());
                    double.TryParse(stdCuttingTimeString, out stdCuttingTime);
                    string stdCuttingTimeUnitString = Convert.ToString(db.tblmasterparts_st_sw.Where(m => m.PartNo == PartNoString && m.OpNo == OperationNoString).Select(m => m.StdCuttingTimeUnit).FirstOrDefault());
                    if (stdCuttingTimeUnitString == "Hrs")
                    {
                        stdCuttingTime = stdCuttingTime * 60;
                    }
                    else if (stdCuttingTimeUnitString == "Sec")
                    {
                        stdCuttingTime = stdCuttingTime / 60;
                    } //else its in minutes.
                    worksheet.Cells["T" + Row].Value = Math.Round(stdCuttingTime, 2); //To Hours

                    //For Summarized and Graph
                    DTAll3.Rows.Add(correctedDate, WorkOrderNo, PartNoString, OperationNoInt, AvgCuttingTime);
                    DataRow dr = DTConsolidatedPL.AsEnumerable().FirstOrDefault(r => r.Field<string>("CorrectedDate") == correctedDate && r.Field<string>("WorkOrderNo") == WorkOrderNo);
                    if (dr == null)
                    {
                        //plant, shop, cell, macINV, WcName, CorrectedDate,WOPF,WOProcessed,
                        //TotalWOQty,TotalTarget,TotalDelivered,TargetNC,TotalValueAdding,TotalLosses,TotalSetUp
                        DTConsolidatedPL.Rows.Add(correctedDate, delQty, delQty, 0, CuttingTime, stdCuttingTime, SettingTime, SelfInsepectionTime, WorkOrderNo);
                    }
                    Row++;
                }

                DataRow dr1 = DTConsolidatedPL.AsEnumerable().FirstOrDefault(r => r.Field<string>("CorrectedDate") == correctedDate);
                if (dr1 != null)
                {
                    //Cellwise Border for Today
                    worksheet.Cells["B" + StartingRowForToday + ":T" + Row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":T" + Row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":T" + Row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["B" + StartingRowForToday + ":T" + Row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    ////Excel:: Border Around Cells.
                    //worksheet.Cells["B" + StartingRowForToday + ":S" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                    //worksheet.Cells["B" + Row + ":S" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    worksheet.Cells["B" + Row + ":G" + Row].Merge = true;
                    worksheet.Cells["B" + Row].Value = "Summarized for ";
                    worksheet.Cells["B" + Row].Style.Font.Bold = true;
                    worksheet.Cells["B" + Row + ":T" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    worksheet.Cells["H" + Row].Value = dr1["CorrectedDate"];
                    string Target = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                        .Sum(x => x.Field<int>("Target")).ToString();
                    int TargetInt = 0;
                    int.TryParse(Target, out TargetInt);
                    worksheet.Cells["M" + Row].Value = TargetInt;

                    worksheet.Cells["H" + Row].Value = dr1["CorrectedDate"];
                    string Delivered = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                        .Sum(x => x.Field<int>("Delivered")).ToString();
                    int DeliveredInt = 0;
                    int.TryParse(Delivered, out DeliveredInt);
                    worksheet.Cells["N" + Row].Value = DeliveredInt;
                    worksheet.Cells["O" + Row].Value = 0; //Rejected Qty

                    string ActualCuttingTime = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                        .Sum(x => x.Field<double>("ActualCutTime")).ToString();
                    double ActualCuttingTimeInt = 0;
                    double.TryParse(ActualCuttingTime, out ActualCuttingTimeInt);
                    worksheet.Cells["R" + Row].Value = Math.Round(ActualCuttingTimeInt, 1);

                    string NCCutTime = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                        .Sum(x => x.Field<double>("NCCutTime")).ToString();
                    double NCCutTimeInt = 0;
                    double.TryParse(NCCutTime, out NCCutTimeInt);
                    worksheet.Cells["T" + Row].Value = Math.Round(NCCutTimeInt, 1);

                    string SetupTime = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                        .Sum(x => x.Field<double>("SetupTime")).ToString();
                    double SetupTimeInt = 0;
                    double.TryParse(SetupTime, out SetupTimeInt);
                    worksheet.Cells["P" + Row].Value = Math.Round(SetupTimeInt, 1);

                    string SelfInspection = DTConsolidatedPL.AsEnumerable().Where(y => y.Field<string>("CorrectedDate") == correctedDate)
                       .Sum(x => x.Field<double>("SelfInspection")).ToString();
                    double SelfInspectionInt = 0;
                    double.TryParse(SelfInspection, out SelfInspectionInt);
                    worksheet.Cells["Q" + Row].Value = Math.Round(SelfInspectionInt, 1);

                    Row++;
                }
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }

            #region OLD Scraped on 2017-03-20
            ////Graphs OverAll
            ////1)Get Distinct WONo's
            ////2)Get Distinct Date's
            ////3) print WCnames in single row
            ////loop dates and inside that loop wono's and push to excel

            //List<string> WONos = (from DataRow dr in DTWONoCuttingTime.Rows
            //                      select (string)dr["WONo"]).Distinct().ToList();

            //List<string> CorrectedDates = (from DataRow dr in DTWONoCuttingTime.Rows
            //                               select (string)dr["CorrectedDate"]).Distinct().ToList();

            //Dictionary<string, string> WCNoXcelCell = new Dictionary<string, string>();
            //int rowGraph = 5;
            //worksheetGraph.Cells["C" + rowGraph].Value = "Date";
            //int col = 4;
            //if (DTWONoCuttingTime.Rows.Count > 0)
            //{
            //    //worksheetGraph.Cells
            //    for (int i = 0; i < WONos.Count(); i++)
            //    {
            //        string ExcelCol = ExcelColumnFromNumber(col);
            //        worksheetGraph.Cells[ExcelCol + rowGraph].Value = WONos[i];
            //        WCNoXcelCell.Add(WONos[i], ExcelCol);
            //        col += 1;
            //    }
            //}
            //rowGraph = 6;

            //// 1st Add 0's to Entire Space,
            //string ForZerosCol = ExcelColumnFromNumber(3 + WONos.Count);
            ////worksheetGraph.Cells["D6:" + ForZerosCol + (5 + CorrectedDates.Count)].Value = "=NA()"; //Data starting from 5

            //for (int i = 0; i < CorrectedDates.Count(); i++)
            //{
            //    for (int j = 0; j < DTWONoCuttingTime.Rows.Count; j++)
            //    {
            //        //worksheetGraph.Cells
            //        //worksheetGraph.Cells["C" + rowGraph++].Value = CorrectedDates[i];
            //        if (DTWONoCuttingTime.Rows[j][0] == CorrectedDates[i])
            //        {

            //            worksheetGraph.Cells["C" + rowGraph].Value = CorrectedDates[i];

            //            string wono = Convert.ToString(DTWONoCuttingTime.Rows[j][1]);
            //            string ColName = null;
            //            if (WCNoXcelCell.TryGetValue(wono, out ColName))
            //            {
            //                double CuttingTime = Convert.ToDouble(DTWONoCuttingTime.Rows[j][2]);
            //                worksheetGraph.Cells[ColName + rowGraph].Value = Math.Round(CuttingTime / 60, 1);
            //            }
            //        }
            //    }
            //    rowGraph++;
            //}

            //ExcelChart chart1 = worksheetGraph.Drawings.AddChart("chart1", eChartType.ColumnClustered3D);
            //var chartIDAndUnID = (ExcelBarChart)chart1;
            //chart1.SetPosition(20, 40);
            //chart1.SetSize((150 * CorrectedDates.Count()), 550);
            //chart1.Title.Font.Bold = true;
            //chart1.Title.Font.Size = 18;
            //chart1.YAxis.MinorTickMark = eAxisTickMark.None;
            //chart1.XAxis.MajorTickMark = eAxisTickMark.None;

            //chart1.Title.Text = "Part Learning";
            //chart1.Style = eChartStyle.Style18;
            //chart1.Legend.Position = eLegendPosition.Bottom;
            ////chart1.d
            //chartIDAndUnID.DataLabel.ShowValue = true;
            //chart1.XAxis.Title.Text = "Completed Dates";
            //chart1.YAxis.Title.Text = "Total Cutting Time";

            //for (var i = 0; i < WONos.Count; i++)
            //{
            //    string headerName = WONos[i];
            //    string ColName = null;
            //    if (WCNoXcelCell.TryGetValue(headerName, out ColName))
            //    {
            //        //var seri = chart.Series.Add("D6:D8", "C6:C" + (5 + CorrectedDates.Count()));
            //        var seri = chart1.Series.Add(ColName + "6:" + ColName + (5 + CorrectedDates.Count()), "C6:C" + (5 + CorrectedDates.Count()));
            //        seri.HeaderAddress = new ExcelAddress("'Graphs'!" + ColName + 5);
            //    }
            //}

            ////Graphs for Individual OperationNo's for Every Date Selected.
            //List<int> OperationNo = (from DataRow dr in DTAll3.Rows
            //                         orderby dr["OpNo"]
            //                         select (int)dr["OpNo"]
            //                         ).Distinct().ToList();

            ////i will need summarized version of DTAll3 DataTable
            //DataTable DTAll3Clone = DTAll3.Clone();
            //int rowPixel = 560, colPixel = 0;
            //double MaxCuttingTime = 0;
            //rowGraph = CorrectedDates.Count + 6 + 3; //keep extra 3 rows.
            //List<double> ForAvg = new List<double>();
            //for (int Opno = 0; Opno < OperationNo.Count; Opno++)
            //{
            //    //push dates into 1st row  of data
            //    //C     D           E           F
            //    //	    2017-01-11	2017-01-12	2017-01-13
            //    //74	343	        463	        34
            //    ForAvg.Clear();
            //    //we need to increment twice to maintain gaps
            //    rowGraph++;
            //    int rowDatesCommon = rowGraph;
            //    //cannnot use same date range col for each opNo so foreach opno print dates
            //    Dictionary<string, string> DateColName = new Dictionary<string, string>();
            //    for (int i = 0; i < CorrectedDates.Count(); i++)
            //    {
            //        string ExcelCol = ExcelColumnFromNumber(4 + i);
            //        worksheetGraph.Cells[ExcelCol + rowGraph].Value = CorrectedDates[i];
            //        DateColName.Add(CorrectedDates[i], ExcelCol);
            //    }
            //    //now in next line print opno and values for each date. So we are incrementing.
            //    rowGraph = rowGraph + 1;

            //    //1st push OPNo into Col 'C' for each OpNo Graph
            //    worksheetGraph.Cells["C" + rowGraph].Value = OperationNo[Opno];
            //    int OpNoInt = OperationNo[Opno];
            //    //for this OpNo push all date values into excel, into respective date columns.

            //    for (int i = 0; i < CorrectedDates.Count(); i++)
            //    {
            //        //now based on date get the sum of cuttingTime and push it into excel

            //        var CuttingValue = 0;
            //        CuttingValue = (from t in DTAll3.AsEnumerable()
            //                        where t["CorrectedDate"].ToString() == CorrectedDates[i]
            //                        && t["OpNo"].ToString() == OpNoInt.ToString()
            //                        select Convert.ToInt32(t["CuttingTime"])).Sum();

            //        //based on date get the column where to push this value
            //        string ColName = null;
            //        if (DateColName.TryGetValue(CorrectedDates[i], out ColName))
            //        {
            //            double cuttingvalLong = Convert.ToDouble(CuttingValue) / 60;
            //            ForAvg.Add(cuttingvalLong);
            //            worksheetGraph.Cells[ColName + rowGraph].Value = Math.Round(cuttingvalLong, 1);
            //            MaxCuttingTime = cuttingvalLong >= MaxCuttingTime ? cuttingvalLong : MaxCuttingTime;
            //        }
            //    }

            //    //to set Position of Graph in Excel
            //    if (Opno != 0)
            //    {
            //        if (Opno % 4 == 0)
            //        {
            //            rowPixel += 410;
            //            colPixel = 0;
            //        }
            //        else
            //        {
            //            colPixel += 410;
            //        }
            //    }

            //    //To Get Average Values.
            //    //yOU will have to get different ABC for each row 
            //    //We Want AA-AB, AC-AD, AE-AF so add opno twice.
            //    string Col1 = ExcelColumnFromNumber(Opno + Opno + 25);
            //    string Col2 = ExcelColumnFromNumber(Opno + Opno + 25 + 1);

            //    for (int i = 0; i < ForAvg.Count; i++)
            //    {
            //        worksheetGraph.Cells[Col1 + (50 + i)].Value = "Avg";
            //        worksheetGraph.Cells[Col2 + (50 + i)].Value = ForAvg.Sum() / ForAvg.Count;
            //    }

            //    ExcelChart chart01 = worksheetGraph.Drawings.AddChart("chart0" + Opno, eChartType.ColumnClustered);
            //    var chartOpNO = (ExcelBarChart)chart01;
            //    chart01.SetPosition(rowPixel, colPixel);
            //    chart01.Title.Font.Bold = true;
            //    chart01.Title.Font.Size = 18;
            //    chart01.YAxis.MinorTickMark = eAxisTickMark.None;
            //    chart01.XAxis.MajorTickMark = eAxisTickMark.None;
            //    chart01.YAxis.MaxValue = (Convert.ToInt32(MaxCuttingTime) + 5);
            //    chart01.YAxis.MinValue = 0;
            //    chart01.Legend.Remove();
            //    //chartOEE.YAxis.MaxValue = 100;
            //    chart01.SetSize(400, 400);
            //    chart01.Title.Text = "Cutting Time for OPNo: " + OperationNo[Opno].ToString();
            //    chart01.Style = eChartStyle.Style18;
            //    //chart01.Legend.Position = eLegendPosition.Bottom;
            //    //chart01.d
            //    chartOpNO.DataLabel.ShowValue = true;

            //    string LastDateCol = ExcelColumnFromNumber(3 + CorrectedDates.Count());
            //    //var seri = chart1.Series.Add(ColName + "6:" + ColName + (5 + CorrectedDates.Count()), "C6:C" + (5 + CorrectedDates.Count()));
            //    //var seri1 = chart01.Series.Add("C" + rowGraph + ":" + LastDateCol + (4 + CorrectedDates.Count()) , "D" + (rowDatesCommon - 1) + ":" +LastDateCol + (4 + CorrectedDates.Count()));
            //    // var seri1 = chart01.Series.Add("D" + rowGraph + ":" + LastDateCol + rowGraph, "C" + (rowDatesCommon) + ":" + LastDateCol + (rowDatesCommon));
            //    //var seri1 = chart01.Series.Add("D" + rowDatesCommon + ":" + LastDateCol + rowDatesCommon, "C" + (rowGraph) + ":" + LastDateCol + (rowGraph));
            //    var seri1 = chart01.Series.Add("D" + (rowGraph) + ":" + LastDateCol + (rowGraph), "D" + rowDatesCommon + ":" + LastDateCol + rowDatesCommon);
            //    var chartTypeL1 = chart01.PlotArea.ChartTypes.Add(eChartType.Line);
            //    var serieL2 = chartTypeL1.Series.Add(worksheetGraph.Cells[Col2 + "50:" + Col2 + (50 + ForAvg.Count - 1)], worksheetGraph.Cells[Col1 + "50:" + Col1 + (50 + ForAvg.Count - 1)]);
            //    rowGraph++;
            //}
            #endregion

            #region OverAll WO wise Stacked OpNo's.
            //Distinct OpNo and Their Column Names.
            //Distinct WONo's.

            var WONoDistinct = (from DataRow row in DTAll3.Rows
                                select row["WONo"]).Distinct();
            var OpNoDistinct = (from DataRow row in DTAll3.Rows
                                select row["OpNo"]).Distinct();

            int rowG = 4, colG = 5;
            int width = 0, height = 0;
            int StartingRow = 0;
            int DistinctOpNoCount = 0;
            if (WONoDistinct.Count() > 0 && OpNoDistinct.Count() > 0)
            {
                //Insert Header Op1,Op2....
                worksheetGraph.Cells["D" + rowG].Value = "WONo";
                Dictionary<int, string> OpNameCol = new Dictionary<int, string>();
                foreach (var opG in OpNoDistinct)
                {
                    string ColName = ExcelColumnFromNumber(colG);
                    int OpNameInt = Convert.ToInt32(opG);
                    worksheetGraph.Cells[ColName + rowG].Value = "Op" + OpNameInt;
                    OpNameCol.Add(OpNameInt, ColName);
                    colG++;
                }

                rowG = 5;
                colG = 4;
                int rowPixel = 20, colPixel = 40, Opno = 0;
                StartingRow = 0;
                foreach (var WoNoG in WONoDistinct)
                {
                    StartingRow = rowG;
                    string WONo = Convert.ToString(WoNoG);
                    bool OpNoIsWritten = false;
                    foreach (var opG in OpNoDistinct)
                    {
                        //now get Value and if exists in DT for This WO and OpNo, get ColName Based on ColValue @ Row = 4.
                        int OpNo = Convert.ToInt32(opG);
                        double CuttingTimeToExcel = 0;
                        CuttingTimeToExcel = DTAll3.AsEnumerable().Where(x => x.Field<string>("WONo") == @WONo && x.Field<int>("OpNo") == @OpNo).Sum(x => x.Field<double>("CuttingTime"));
                        if (CuttingTimeToExcel != 0 && OpNameCol.ContainsKey(OpNo))
                        {
                            //get the Max of Date for this WONo
                            string CDate = DTAll3.AsEnumerable().Where(x => x.Field<string>("WONo") == @WONo).Max(x => x.Field<string>("CorrectedDate"));
                            worksheetGraph.Cells["D" + rowG].Value = WONo + " & " + CDate;
                            string ColName = OpNameCol[OpNo];
                            worksheetGraph.Cells[ColName + rowG].Value = CuttingTimeToExcel;
                            OpNoIsWritten = true;
                        }
                    }
                    if (OpNoIsWritten)
                    {
                        rowG++;
                    }
                }

                //Now Plot the Graph.
                width = 0; height = 0;
                if (WONoDistinct.Count() > 3)
                {
                    width = (WONoDistinct.Count() * 100 + 30);
                }
                else
                {
                    width = 350;
                }
                height = 350;

                ExcelChart chart01 = worksheetGraph.Drawings.AddChart("OverAll WOs", eChartType.ColumnStacked);
                var chartOpNO = (ExcelBarChart)chart01;
                chart01.SetPosition(rowPixel, colPixel);
                chart01.Title.Font.Bold = true;
                chart01.Title.Font.Size = 18;
                chart01.YAxis.MinorTickMark = eAxisTickMark.None;
                chart01.XAxis.MajorTickMark = eAxisTickMark.None;
                chart01.YAxis.MinValue = 0.1;
                //chart01.Legend.Remove();
                chart01.Legend.Position = eLegendPosition.Bottom;
                //chartOEE.YAxis.MaxValue = 100;
                chart01.SetSize(width, height); //width , height
                chart01.Title.Text = "OverAll Work Order wise";
                chart01.Style = eChartStyle.Style18;

                //chart01.d
                chartOpNO.DataLabel.ShowValue = true;

                int WOLooperInt = 5;
                DistinctOpNoCount = OpNoDistinct.Count();
                string finalOpNoColName = ExcelColumnFromNumber(4 + DistinctOpNoCount); //OpNoStarts @ Col 5, so add 4 to Column Count
                var WORange = worksheetGraph.Cells["D5:D" + (4 + WONoDistinct.Count())];
                foreach (var WoNoG in OpNoDistinct)
                {
                    int WONo = Convert.ToInt32(WoNoG);
                    string ColName = OpNameCol[WONo];
                    var ran1 = worksheetGraph.Cells[ColName + WOLooperInt + ":" + ColName + (4 + WONoDistinct.Count())];
                    var serie1 = chart01.Series.Add(ran1, WORange);
                    serie1.HeaderAddress = worksheetGraph.Cells[ColName + "4"];
                }

                chart01.DisplayBlanksAs = eDisplayBlanksAs.Gap; //To Make sure graphs don't plot 0. Awesome. i figured it out.
            }
            #endregion
            rowG += 5;
            int Row1st = rowG; colG = 5;
            #region OpNo wise Graphs
            if (WONoDistinct.Count() > 3)
            {
                width = (WONoDistinct.Count() * 100 + 30);
            }
            else
            {
                width = 400;
            }
            height = 350;
            Dictionary<string, string> WONameCol = new Dictionary<string, string>();
            foreach (var WoNoG in WONoDistinct)
            {
                string ColName = ExcelColumnFromNumber(colG);
                string WONo = Convert.ToString(WoNoG);
                worksheetGraph.Cells[ColName + rowG].Value = WONo;
                WONameCol.Add(WONo, ColName);
                colG++;
            }
            rowG++;

            int PixelLeftInt = 10, PixelTopInt = 400;
            int OpNoLooper = 1;
            foreach (var opG in OpNoDistinct)
            {
                //to set Position of Graph in Excel
                int OpNo = Convert.ToInt32(opG);
                if (OpNo != 0 && OpNoLooper != 1)
                {
                    if (OpNoLooper % 3 == 0)
                    {
                        PixelTopInt += 400;
                        PixelLeftInt = 0;
                    }
                    else
                    {
                        PixelLeftInt += 400;
                    }
                }


                foreach (var WoNoG in WONoDistinct)
                {
                    StartingRow = rowG;
                    string WONo = Convert.ToString(WoNoG);

                    double CuttingTimeToExcel = 0;
                    CuttingTimeToExcel = DTAll3.AsEnumerable().Where(x => x.Field<string>("WONo") == @WONo && x.Field<int>("OpNo") == @OpNo).Sum(x => x.Field<double>("CuttingTime"));
                    if (CuttingTimeToExcel != 0 && WONameCol.ContainsKey(WONo))
                    {
                        worksheetGraph.Cells["D" + rowG].Value = OpNo;
                        string ColName = WONameCol[WONo];
                        worksheetGraph.Cells[ColName + rowG].Value = CuttingTimeToExcel;
                    }
                }

                string CellsOfTop5LossColNames = null;
                string CellsOfTop5LossColValues = null;

                //push and Set Cell Range for Graphs.
                string finalWONoColName = ExcelColumnFromNumber(4 + DistinctOpNoCount);
                ExcelRange erLossesRangechartTop5LossesvALUE = worksheetGraph.Cells["E" + rowG + ":" + finalWONoColName + rowG];
                ExcelRange erLossesRangechartTop5LossesNAMES = worksheetGraph.Cells["E" + Row1st + ":" + (finalWONoColName + Row1st)];

                ExcelChart chartTop5Losses1 = (ExcelBarChart)worksheetGraph.Drawings.AddChart("barChartTop5Losses0" + OpNoLooper, eChartType.ColumnClustered);
                var chartTop5Losses = (ExcelBarChart)chartTop5Losses1;
                chartTop5Losses.SetSize(390, 300);
                chartTop5Losses.SetPosition(PixelTopInt, PixelLeftInt);
                //string blah = "CY11,CZ11,DA11,DC11,DD11"; //This Works 
                //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells[blah];
                //ExcelRange erLossesRangechartTop5LossesvALUE = worksheet.Cells["CY11,CZ11,DA11,DC11,DD11"];
                //ExcelRange erLossesRangechartTop5LossesNAMES = worksheet.Cells["CY3,CZ3,DA3,DC3,DD3"];

                chartTop5Losses.Title.Text = " OpNo. " + Convert.ToString(opG);
                chartTop5Losses.Style = eChartStyle.Style19;
                chartTop5Losses.Legend.Remove();
                //chartTop5Losses.Legend.Position = eLegendPosition.Bottom;
                chartTop5Losses.DataLabel.ShowValue = true;
                chartTop5Losses.XAxis.Font.Size = 8;
                chartTop5Losses.YAxis.MinorTickMark = eAxisTickMark.None;
                chartTop5Losses.XAxis.MajorTickMark = eAxisTickMark.None;
                //chartTop5Losses.Legend.Add();
                chartTop5Losses.Series.Add(erLossesRangechartTop5LossesvALUE, erLossesRangechartTop5LossesNAMES);
                RemoveGridLines(ref chartTop5Losses1);

                //chartTop5Losses1.DisplayBlanksAs = eDisplayBlanksAs.Gap; //To Make sure graphs don't plot 0. Awesome. i figured it out.
                rowG++;
                OpNoLooper++;
            }

            #endregion

            #region Save and Download

            //Hide Values
            //Color ColorHexWhite = System.Drawing.Color.White;
            //worksheetGraph.Cells["A1:Z50"].Style.Font.Color.SetColor(ColorHexWhite);

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            worksheetGraph.View.ShowGridLines = false;
            p.Save();

            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "PartLearning" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "PartLearning" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            #endregion
        }

        List<string> GetHierarchyData(int MachineID)
        {
            List<string> HierarchyData = new List<string>();
            //1st get PlantName or -
            //2nd get ShopName or -
            //3rd get CellName or -
            //4th get MachineName.

            using (i_facility_talEntities1 dbMac = new i_facility_talEntities1())
            {
                var machineData = dbMac.tblmachinedetails.Where(m => m.MachineID == MachineID).FirstOrDefault();
                int PlantID = Convert.ToInt32(machineData.PlantID);
                string name = "-";
                name = dbMac.tblplants.Where(m => m.PlantID == PlantID).Select(m => m.PlantName).FirstOrDefault();
                HierarchyData.Add(name);

                string ShopIDString = Convert.ToString(machineData.ShopID);
                int value;
                if (int.TryParse(ShopIDString, out value))
                {
                    name = dbMac.tblshops.Where(m => m.ShopID == value).Select(m => m.ShopName).FirstOrDefault();
                    HierarchyData.Add(name.ToString());
                }
                else
                {
                    HierarchyData.Add("-");
                }

                string CellIDString = Convert.ToString(machineData.CellID);
                if (int.TryParse(CellIDString, out value))
                {
                    name = dbMac.tblcells.Where(m => m.CellID == value).Select(m => m.CellName).FirstOrDefault();
                    HierarchyData.Add(name.ToString());
                }
                else
                {
                    HierarchyData.Add("-");
                }
                HierarchyData.Add(Convert.ToString(machineData.MachineInvNo));
                HierarchyData.Add(Convert.ToString(machineData.MachineDispName));
            }
            return HierarchyData;
        }

        List<string> ExtractPartNo(string PartsList)
        {
            List<string> parts = new List<string>();
            if (PartsList.Contains(','))
            {
                string[] mails = PartsList.Split(',');
                parts.AddRange(mails);
                parts = parts.Where(s => !string.IsNullOrWhiteSpace(s)).Where(s => !string.IsNullOrEmpty(s)).Distinct().ToList();
            }

            return parts;
        }
        string ExtractPartNoStringRegex(string PartsList)
        {
            //Old Format : 'FW81184','KH42796','KH46372'
            //New Format(2017-03-18) : '92200-02311-130(.*)|92209-02402-101(.*)|92308-02105-120(.*)'
            string parts = null;
            if (PartsList.Contains(','))
            {
                string[] PartsArray = PartsList.Split(',');
                for (int j = 0; j < PartsArray.Count(); j++)
                {
                    if (!string.IsNullOrEmpty(PartsArray[j]))
                    {
                        if (j == 0)
                        {
                            parts += "'" + PartsArray[j]; //Start With
                        }
                        else
                        {
                            //parts += "'" + "," + "'" + PartsArray[j];
                            parts += @"(.*)" + "|" + PartsArray[j];
                        }
                    }
                }
                parts += @"(.*)" + "'"; //End With
            }
            return parts;
        }

        public async Task<double> AnyLossDuration(string StartDate, string EndDate, string LossName, int MachineID)
        {
            double duration = 0;


            return duration;
        }
        public async Task<double> GetActualCuttingTime(int machineID, string PartNo, string WONo, string OpNo, string UsedDateForExcel)
        {
            double ACTime = 0;
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                //var HMIData = dbhmi.tblhmiscreens.Where(m => m.MachineID == machineID && m.PartNo == PartNo && m.Work_Order_No == WONo && m.OperationNo == OpNo && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
                //foreach (var row in HMIData)
                //{
                //    DateTime startTime = Convert.ToDateTime(row.Date);
                //    //if(Convert.ToString(row.Time)) //Check for Null or Empty
                //    DateTime endTime = Convert.ToDateTime(row.Time);

                //    ACTime += GetGreen(UsedDateForExcel, startTime, endTime, machineID);
                //}

                //2017-03-17
            }
            return ACTime;
        }

        //UnIdentified Report Method
        public async Task<string> UnIdentifiedReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {
            string lowestLevelName = null;
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\UnIdentifiedLossReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                            lowestLevelName = db.tblplants.Where(m => m.PlantID == plantId).Select(m => m.PlantName).FirstOrDefault();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                        lowestLevelName = db.tblshops.Where(m => m.ShopID == shopId).Select(m => m.ShopName).FirstOrDefault();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                    lowestLevelName = db.tblcells.Where(m => m.CellID == cellId).Select(m => m.CellName).FirstOrDefault();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                lowestLevelName = db.tblmachinedetails.Where(m => m.MachineID == wcId).Select(m => m.MachineDispName).FirstOrDefault();
                MacCount = 1;
            }

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "UnIdentifiedLossReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "UnIdentifiedLossReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            //ExcelWorksheet worksheetOA = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                //worksheetOA = p.Workbook.Worksheets.Add("Summarized " + System.DateTime.Now.ToString("dd-MM-yyyy"), TemplateSummarized);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                //worksheetOA = p.Workbook.Worksheets.Add("Graphs " + System.DateTime.Now.ToString("dd-MM-yyyy"), TemplateSummarized);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            //Step1: Get the Summarized according to Name. for each WC from StartDate - EndDate
            //( Leave that many rows blank and Fill WorkCenter wise 1st.
            

            //DataTable for Consolidated Duration 
            DataTable DTNoLogin = new DataTable();
            DTNoLogin.Columns.Add("MachineID", typeof(int));
            DTNoLogin.Columns.Add("Duration", typeof(int));


            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5;
            int Sno = 1;
            for (int i = 0; i < TotalDay + 1; i++)
            {
                DateTime endDateTime = Convert.ToDateTime(UsedDateForExcel.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
                string startDateTime = UsedDateForExcel.ToString("yyyy-MM-dd");

                DataTable HMIData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = null;
                    if (lowestLevel == "Plant")
                    {
                        //query = "select  MachineID, StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) order by MachineID,StartDateTime; ";
                        query = "select  MachineID, StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate	= '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where PlantId = " + plantId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) order by MachineID,StartDateTime; ";
                    }
                    else if (lowestLevel == "Shop")
                    {
                        //query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) order by MachineID,StartDateTime; ";
                        query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where ShopID = " + shopId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) order by MachineID,StartDateTime; ";
                    }
                    else if (lowestLevel == "Cell")
                    {
                        //query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) order by MachineID,StartDateTime; ";
                        query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "'  and  MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where CellID = " + cellId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) order by MachineID,StartDateTime; ";
                    }
                    else if (lowestLevel == "WorkCentre")
                    {
                        //query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ) order by MachineID,StartDateTime; ";
                        query = "select  MachineID,  StartDateTime, EndDateTime from i_facility_tal.dbo.tbllossofentry where CorrectedDate = '" + UsedDateForExcel.ToString("yyyy-MM-dd") + "' and MessageCodeID = 999 and DoneWithRow = 1 and MachineID IN (select MachineID from i_facility_tal.dbo.tblmachinedetails where MachineID = " + wcId + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ) order by MachineID,StartDateTime; ";
                    }
                    SqlDataAdapter da = new SqlDataAdapter(query, mc.msqlConnection);
                    da.Fill(HMIData);
                    mc.close();
                }
                double individualDateDuration = 0;
                int thisDatesDataStartsAt = Row;

                for (int n = 0; n < HMIData.Rows.Count; n++)
                {
                    if (n == 0 && i != 0)
                    {
                        Row++;
                        thisDatesDataStartsAt = Row;
                    }
                    int MachineID = Convert.ToInt32(HMIData.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);

                    worksheet.Cells["B" + Row].Value = Sno++;
                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                    worksheet.Cells["F" + Row].Value = HierarchyData[3];

                    //string Duration = null;
                    double DurInSeconds = 0;
                    if (!string.IsNullOrEmpty(Convert.ToString(HMIData.Rows[n][1])) && !string.IsNullOrEmpty(Convert.ToString(HMIData.Rows[n][1])))
                    {
                        DateTime StartTime = Convert.ToDateTime(HMIData.Rows[n][1]);
                        DateTime EndTime = Convert.ToDateTime(HMIData.Rows[n][2]);
                        DurInSeconds = Convert.ToDouble(EndTime.Subtract(StartTime).TotalSeconds);
                    }
                    individualDateDuration += Convert.ToDouble(DurInSeconds);

                    worksheet.Cells["G" + Row].Value = Convert.ToDateTime(HMIData.Rows[n][1]).ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cells["H" + Row].Value = Convert.ToDateTime(HMIData.Rows[n][2]).ToString("yyyy-MM-dd HH:mm:ss");

                    worksheet.Cells["I" + Row].Value = Math.Round((DurInSeconds / 60), 2);
                    DataRow dr = DTNoLogin.Select("MachineID = " + MachineID).FirstOrDefault(); // finds all rows with id==2 and selects first or null if haven't found any
                    if (dr != null)
                    {
                        Int64 DurationPrev = Convert.ToInt64(dr["Duration"]); //get lossduration and update it.
                        dr["Duration"] = (DurationPrev + DurInSeconds);
                    }
                    else
                    {
                        DTNoLogin.Rows.Add(MachineID, DurInSeconds);
                    }
                    Row++;
                }

                if (Row > thisDatesDataStartsAt)
                {
                    //thin border for all cells
                    var modelTable = worksheet.Cells["B" + thisDatesDataStartsAt + ":I" + Row];
                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    //Insert OverAll NoLogin Time
                    worksheet.Cells["B" + Row + ":H" + Row].Merge = true;
                    worksheet.Cells["B" + Row].Value = "Duration";
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Font.Bold = true;
                    worksheet.Cells["B" + Row + ":I" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                    double SumNoLogin = 0;
                    SumNoLogin = Convert.ToDouble(individualDateDuration / 60);
                    double RoundOffSumNoLogin = Math.Round(SumNoLogin, 1);
                    worksheet.Cells["I" + Row].Value = RoundOffSumNoLogin;

                    //Excel:: Border Around Cells.
                    worksheet.Cells["B" + thisDatesDataStartsAt + ":I" + "" + Row].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                }
                Row++;
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);
            }
            //to Align Center for sheet detailed
            worksheet.Cells["B" + 5 + ":I" + Row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "UnIdentifiedLossReport" + lowestLevelName.Trim() + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "UnIdentifiedLossReport " + lowestLevelName + " " + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            return path1;
        }

        public async void StdTimeWeightReportExcel(string StartDate, string EndDate, string PlantID, string ShopID = null, string CellID = null, string WorkCenterID = null)
        {

            #region MacCount & LowestLevel
            string lowestLevel = null;
            int MacCount = 0;
            int plantId = 0, shopId = 0, cellId = 0, wcId = 0;
            if (string.IsNullOrEmpty(WorkCenterID))
            {
                if (string.IsNullOrEmpty(CellID))
                {
                    if (string.IsNullOrEmpty(ShopID))
                    {
                        if (string.IsNullOrEmpty(PlantID))
                        {
                            //donothing
                        }
                        else
                        {
                            lowestLevel = "Plant";
                            plantId = Convert.ToInt32(PlantID);
                            MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.PlantID == plantId).ToList().Count();
                        }
                    }
                    else
                    {
                        lowestLevel = "Shop";
                        shopId = Convert.ToInt32(ShopID);
                        MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.ShopID == shopId).ToList().Count();
                    }
                }
                else
                {
                    lowestLevel = "Cell";
                    cellId = Convert.ToInt32(CellID);
                    MacCount = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.CellID == cellId).ToList().Count();
                }
            }
            else
            {
                lowestLevel = "WorkCentre";
                wcId = Convert.ToInt32(WorkCenterID);
                MacCount = 1;
            }

            #endregion

            #region Excel and Stuff

            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(StartDate) == true)
            {
                StartDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndDate) == true)
            {
                EndDate = StartDate;
            }

            DateTime frmDate = Convert.ToDateTime(StartDate);
            DateTime toDate = Convert.ToDateTime(EndDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\NewTemplates\MasterPartsSTD.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];
            ExcelWorksheet TemplateGraph = templatep.Workbook.Worksheets[2];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            bool exists = System.IO.Directory.Exists(FileDir);
            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MasterData" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MasterData" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;
            ExcelWorksheet worksheetUpload = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                worksheetUpload = p.Workbook.Worksheets.Add("ToUpload", TemplateGraph);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                 }
                catch (Exception e)
                { }
            }

            if (worksheetUpload == null)
            {
                try{
                    worksheetUpload = p.Workbook.Worksheets.Add("ToUpload1", TemplateGraph);
                }
                catch (Exception e)
                { }
            }
            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);
            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            #endregion

            

            #region Get Machines List
            DataTable machin = new DataTable();
            DateTime endDateTime = Convert.ToDateTime(toDate.AddDays(1).ToString("yyyy-MM-dd") + " " + new TimeSpan(6, 0, 0));
            string startDateTime = frmDate.ToString("yyyy-MM-dd");
            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String query1 = null;
            if (lowestLevel == "Plant")
            {
                //query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT  distinct MachineID FROM i_facility_tal.dbo.tblmachinedetails WHERE PlantID = " + PlantID + " and ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "Shop")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE ShopID = " + ShopID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "Cell")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + " and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE CellID = " + CellID + " and   ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            else if (lowestLevel == "WorkCentre")
            {
                //query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   ( case when (IsDeleted = 1) then ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) end) ) ;";
                query1 = " SELECT * FROM i_facility_tal.dbo.tblmachinedetails WHERE MachineID = " + WorkCenterID + " and  ((InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and IsDeleted = 0) or   (  (IsDeleted = 1) and ( InsertedOn <= '" + endDateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and DeletedDate >= '" + startDateTime + "' ) ) ) ;";
            }
            SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
            da1.Fill(machin);
            mc.close();
            #endregion

            //List <Tuple<string, string, string>> WOPartOPList = new List<Tuple<string, string, string>>();
            Dictionary<string, string> PartOPList = new Dictionary<string, string>();

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);
            //For each Date ...... for all Machines.
            var Col = 'B';
            int Row = 5; // Gap to Insert OverAll data. DataStartRow + MachinesCount + 2(1 for HighestLevel & another for Gap).
            int Sno = 1;
            string finalLossCol = null;

            for (int i = 0; i < TotalDay + 1; i++)
            {
                int StartingRowForToday = Row;
                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");

                int NumMacsToExcel = 0;
                for (int n = 0; n < machin.Rows.Count; n++)
                {
                    NumMacsToExcel++;
                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                    List<string> HierarchyData = GetHierarchyData(MachineID);
                    //Added this machineDetails into Datatable
                    string WCInvNoString = HierarchyData[3];
                    string correctedDate = UsedDateForExcel.ToString("yyyy-MM-dd");

                    #region Get HMI DATA and Push For this machine and Date.
                    //Get general hmi data, later get prodfai && operator Combination based data inside loop

                    var HMIDATA = db.tblworeports.Where(m => m.MachineID == MachineID && m.CorrectedDate == correctedDate).GroupBy(m => m.HMIID).Select(m => m.FirstOrDefault()).ToList();
                    if (HMIDATA != null)
                    {
                        //Now Loop through and Push Data into Excel
                        foreach (var row in HMIDATA)
                        {
                            if (Convert.ToInt32(row.IsMultiWO) == 1) //Its a MultiWorkOrder
                            {
                                #region
                                int HmiID = Convert.ToInt32(row.HMIID);
                                var MulitWOData = db.tblworeports.Where(m => m.HMIID == HmiID).ToList();
                                foreach (var MulitWORow in MulitWOData)
                                {
                                    if (Convert.ToDouble(MulitWORow.NCCuttingTimePerPart) == 0)
                                    {
                                        worksheet.Cells["B" + Row].Value = Sno++;
                                        worksheet.Cells["C" + Row].Value = HierarchyData[0];
                                        worksheet.Cells["D" + Row].Value = HierarchyData[1];
                                        worksheet.Cells["E" + Row].Value = HierarchyData[2];
                                        worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                                        worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV
                                        worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                                        worksheet.Cells["I" + Row].Value = row.Shift;
                                        worksheet.Cells["J" + Row].Value = MulitWORow.PartNo;
                                        worksheet.Cells["K" + Row].Value = MulitWORow.WorkOrderNo;
                                        worksheet.Cells["L" + Row].Value = MulitWORow.OpNo;
                                        string PartNoS = MulitWORow.PartNo;
                                        string WorkOrderNoS = MulitWORow.WorkOrderNo;
                                        string OpNoS = MulitWORow.OpNo;

                                        try
                                        {
                                            Row++;
                                            PartOPList.Add(PartNoS, OpNoS);
                                        }
                                        catch (Exception e)
                                        {
                                        }

                                        string TargetQtys = Convert.ToString(db.tbl_multiwoselection.Where(m => m.HMIID == HmiID && m.PartNo == PartNoS && m.WorkOrder == WorkOrderNoS && m.OperationNo == OpNoS).Select(m => m.TargetQty).FirstOrDefault());
                                        int TargetQty = 0;
                                        int.TryParse(TargetQtys, out TargetQty);
                                        worksheet.Cells["M" + Row].Value = TargetQty;

                                        int delQty = 0;
                                        string Delivered_QtyS = Convert.ToString(db.tbl_multiwoselection.Where(m => m.HMIID == HmiID && m.PartNo == PartNoS && m.WorkOrder == WorkOrderNoS && m.OperationNo == OpNoS).Select(m => m.DeliveredQty).FirstOrDefault());
                                        int.TryParse(Convert.ToString(Delivered_QtyS), out delQty);
                                        worksheet.Cells["N" + Row].Value = delQty;

                                        if (row.IsPF == 0)
                                        {
                                            worksheet.Cells["O" + Row].Value = "Yes";
                                        }
                                        else
                                        {
                                            worksheet.Cells["P" + Row].Value = "Yes";
                                        }
                                        double SettingTime = Convert.ToDouble(row.SettingTime);
                                        worksheet.Cells["Q" + Row].Value = Math.Round(SettingTime, 1);
                                        double Green = Convert.ToDouble(row.CuttingTime);
                                        worksheet.Cells["R" + Row].Value = Green;
                                        double idleTime = Convert.ToDouble(row.Idle - row.SettingTime);
                                        worksheet.Cells["S" + Row].Value = Math.Round(idleTime, 2);
                                        worksheet.Cells["T" + Row].Formula = "=SUM(Q" + Row + ",R" + Row + ",S" + Row + ",U" + Row + ")";
                                        double BreakdownDuration = Convert.ToDouble(row.Breakdown);
                                        worksheet.Cells["U" + Row].Value = Math.Round(BreakdownDuration, 1);
                                        worksheet.Cells["V" + Row].Value = 0;
                                        worksheet.Cells["W" + Row].Value = "0";
                                        worksheet.Cells["X" + Row].Value = row.OperatorName;
                                        worksheet.Cells["Y" + Row].Value = row.Type;
                                        worksheet.Cells["Z" + Row].Value = row.PartNo + "#" + row.OpNo;
                                        double stdCuttingTime = 0;
                                        double ProdOfstdCuttingTimeDelivQty = 0;
                                        stdCuttingTime = Convert.ToDouble(MulitWORow.NCCuttingTimePerPart);
                                        ProdOfstdCuttingTimeDelivQty = delQty * stdCuttingTime;
                                        worksheet.Cells["AA" + Row].Value = stdCuttingTime;
                                        worksheet.Cells["AB" + Row].Value = ProdOfstdCuttingTimeDelivQty;
                                    }
                                }
                                #endregion

                            }
                            else
                            {
                                if (Convert.ToDouble(row.NCCuttingTimePerPart) == 0)
                                {

                                    // Unnecessary WOPartOPList.Add(new Tuple<string, string, string>(Convert.ToString(row.WorkOrderNo), Convert.ToString(row.PartNo), Convert.ToString(row.OpNo)));
                                    try
                                    {
                                        Row++;
                                        PartOPList.Add(row.PartNo, row.OpNo);
                                    }
                                    catch (Exception e)
                                    {
                                    }

                                    #region To push to excel. Single WO.
                                    worksheet.Cells["B" + Row].Value = Sno++;
                                    worksheet.Cells["C" + Row].Value = HierarchyData[0];
                                    worksheet.Cells["D" + Row].Value = HierarchyData[1];
                                    worksheet.Cells["E" + Row].Value = HierarchyData[2];
                                    worksheet.Cells["F" + Row].Value = HierarchyData[4];//Display Name
                                    worksheet.Cells["G" + Row].Value = HierarchyData[3];//Mac INV

                                    worksheet.Cells["H" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                                    worksheet.Cells["I" + Row].Value = row.Shift;
                                    worksheet.Cells["J" + Row].Value = row.PartNo;
                                    worksheet.Cells["K" + Row].Value = row.WorkOrderNo;
                                    worksheet.Cells["L" + Row].Value = row.OpNo;
                                    worksheet.Cells["M" + Row].Value = row.TargetQty;
                                    double delQty = Convert.ToDouble(row.DeliveredQty);
                                    worksheet.Cells["N" + Row].Value = delQty;
                                    if (row.IsPF == 0)
                                    {
                                        worksheet.Cells["O" + Row].Value = "Yes";
                                    }
                                    else
                                    {
                                        worksheet.Cells["P" + Row].Value = "Yes";
                                    }

                                    double SettingTime = Convert.ToDouble(row.SettingTime);
                                    worksheet.Cells["Q" + Row].Value = Math.Round(SettingTime, 1);

                                    double Green = Convert.ToDouble(row.CuttingTime);
                                    worksheet.Cells["R" + Row].Value = Green;

                                    double idleTime = Convert.ToDouble(row.Idle - row.SettingTime);
                                    worksheet.Cells["S" + Row].Value = Math.Round(idleTime, 2);

                                    double BreakdownDuration = Convert.ToDouble(row.Breakdown);
                                    worksheet.Cells["U" + Row].Value = Math.Round(BreakdownDuration, 1);

                                    // worksheet.Cells["T" + Row].Formula = "=SUM(Q" + Row + ",R" + Row + ",S" + Row + ",T" + Row + ",X" + Row + ")";
                                    worksheet.Cells["T" + Row].Formula = "=SUM(Q" + Row + ",R" + Row + ",S" + Row + ",U" + Row + ")";
                                    worksheet.Cells["V" + Row].Value = 0;
                                    worksheet.Cells["W" + Row].Value = "0";
                                    worksheet.Cells["X" + Row].Value = row.OperatorName;

                                    worksheet.Cells["Y" + Row].Value = row.Type;

                                    //To skip a Column Just Increment the ColIndex extra +1
                                    worksheet.Cells["Z" + Row].Value = row.PartNo + "#" + row.OpNo;

                                    double stdCuttingTime = 0;
                                    double ProdOfstdCuttingTimeDelivQty = 0;

                                    stdCuttingTime = Convert.ToDouble(row.NCCuttingTimePerPart);
                                    ProdOfstdCuttingTimeDelivQty = delQty * stdCuttingTime;

                                    worksheet.Cells["AA" + Row].Value = stdCuttingTime;
                                    worksheet.Cells["AB" + Row].Value = ProdOfstdCuttingTimeDelivQty;

                                    #endregion
                                }
                            }

                        }//end of 1 HMIDATA Row
                    }
                    #endregion

                }//End of For Each Machine Loop
                UsedDateForExcel = UsedDateForExcel.AddDays(+1);

            } //End of All day's Loop

            var PartsList = PartOPList.Count();
            int UploadLooper = 3;
            foreach (var row in PartOPList)
            {
                string partno = row.Key;
                string opno = row.Value;

                var ToUploadData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.PartNo.Equals(partno, StringComparison.OrdinalIgnoreCase) && m.OpNo.Equals(opno, StringComparison.OrdinalIgnoreCase)).ToList();

                foreach (var rowInner in ToUploadData)
                {
                    worksheetUpload.Cells["A" + UploadLooper].Value = rowInner.PartNo;
                    worksheetUpload.Cells["B" + UploadLooper].Value = rowInner.OpNo;
                    worksheetUpload.Cells["C" + UploadLooper].Value = rowInner.StdSetupTime;
                    worksheetUpload.Cells["D" + UploadLooper].Value = rowInner.StdSetupTimeUnit; ;
                    worksheetUpload.Cells["E" + UploadLooper].Value = rowInner.StdCuttingTime;
                    worksheetUpload.Cells["F" + UploadLooper].Value = rowInner.StdCuttingTimeUnit;
                    worksheetUpload.Cells["G" + UploadLooper].Value = rowInner.StdChangeoverTime;
                    worksheetUpload.Cells["H" + UploadLooper].Value = rowInner.StdChangeoverTimeUnit;
                    worksheetUpload.Cells["I" + UploadLooper].Value = rowInner.InputWeight;
                    worksheetUpload.Cells["J" + UploadLooper].Value = rowInner.InputWeightUnit;
                    worksheetUpload.Cells["K" + UploadLooper].Value = rowInner.OutputWeight;
                    worksheetUpload.Cells["L" + UploadLooper].Value = rowInner.OutputWeightUnit;
                    worksheetUpload.Cells["M" + UploadLooper].Value = rowInner.MaterialRemovedQty;
                    worksheetUpload.Cells["N" + UploadLooper].Value = rowInner.MaterialRemovedQtyUnit;

                    UploadLooper++;
                }
                if (ToUploadData.Count == 0)
                {
                    worksheetUpload.Cells["A" + UploadLooper].Value = partno;
                    worksheetUpload.Cells["B" + UploadLooper].Value = opno;
                    worksheetUpload.Cells["C" + UploadLooper].Value = 0.00;
                    worksheetUpload.Cells["D" + UploadLooper].Value = "Min";
                    worksheetUpload.Cells["E" + UploadLooper].Value = 0.00;
                    worksheetUpload.Cells["F" + UploadLooper].Value = "Min";
                    worksheetUpload.Cells["G" + UploadLooper].Value = 0.00;
                    worksheetUpload.Cells["H" + UploadLooper].Value = "Min";
                    worksheetUpload.Cells["I" + UploadLooper].Value = 0;
                    worksheetUpload.Cells["J" + UploadLooper].Value = "Kg";
                    worksheetUpload.Cells["K" + UploadLooper].Value = 0;
                    worksheetUpload.Cells["L" + UploadLooper].Value = "Kg";
                    worksheetUpload.Cells["M" + UploadLooper].Value = 0;
                    worksheetUpload.Cells["N" + UploadLooper].Value = "Kg";
                    UploadLooper++;
                }

            }
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            #region Save and Download
            p.Save();

            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "MasterData" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "MasterData" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            #endregion

        }

        //Break Down Report
        public async void generateBreakDownReportExcel(string startDate, string EndtDate)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\BreakDownReport.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "BreakDownReport" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "BreakDownReport" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");
            //var data = db.tblbreakdowns.Where(m => m.CorrectedDate >= FDate && m.CorrectedDate <= TDate);

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,StartTime,EndTime,BreakDownCode,CorrectedDate,Shift FROM [i_facility_tal].[dbo].tblbreakdown WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' AND DoneWithRow = 1 ORDER BY BreakdownID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();
            int bdid = 0;

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 8;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(m => m.MachineID == MachineID).FirstOrDefault();
                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineDispName;
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false)
                    {
                        DateTime startdate = Convert.ToDateTime(dataHolder.Rows[i][1]);
                        worksheet.Cells["D" + Row].Value = startdate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        DateTime Enddate = Convert.ToDateTime(dataHolder.Rows[i][2]);
                        worksheet.Cells["E" + Row].Value = Enddate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][3].ToString()) == false)
                        bdid = Convert.ToInt32(dataHolder.Rows[i][3]);
                    mc.open();
                    String sql2 = "SELECT * FROM [i_facility_tal].[dbo].tblbreakdowncodes WHERE BreakDownCodeID= " + bdid + "";
                    SqlDataAdapter da2 = new SqlDataAdapter(sql2, mc.msqlConnection);
                    DataTable dataHolder2 = new DataTable();
                    da2.Fill(dataHolder2);
                    mc.close();

                    int level = Convert.ToInt32(dataHolder2.Rows[0][4]);
                    if (level == 1)
                    {
                        worksheet.Cells["F" + Row].Value = dataHolder2.Rows[0][1].ToString();
                    }
                    else if (level == 2)
                    {
                        int id = Convert.ToInt32(dataHolder2.Rows[0][5]);
                        var data = db.tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeID == id).SingleOrDefault();
                        worksheet.Cells["G" + Row].Value = dataHolder2.Rows[0][1].ToString();
                        worksheet.Cells["F" + Row].Value = data.LossCode;
                    }
                    else if (level == 3)
                    {
                        int id = Convert.ToInt32(dataHolder2.Rows[0][5]);
                        var data = db.tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeID == id).SingleOrDefault();

                        int id1 = Convert.ToInt32(dataHolder2.Rows[0][6]);
                        var data1 = db.tbllossescodes.Where(m => m.IsDeleted == 0 && m.LossCodeID == id).SingleOrDefault();

                        worksheet.Cells["H" + Row].Value = dataHolder2.Rows[0][1].ToString();
                        worksheet.Cells["F" + Row].Value = data.LossCode;
                        worksheet.Cells["G" + Row].Value = data1.LossCode;
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false && string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        DateTime StartTime = DateTime.Now;
                        StartTime = Convert.ToDateTime(dataHolder.Rows[i][1]);
                        DateTime EndTime = DateTime.Now;
                        EndTime = Convert.ToDateTime(dataHolder.Rows[i][2]);
                        TimeSpan ts = EndTime.Subtract(StartTime);
                        int H = ts.Hours;
                        int M = ts.Minutes;
                        int S = ts.Seconds;
                        string Hs = null, Ms = null, Ss = null;
                        if (H < 10)
                        {
                            Hs = "0" + H;
                        }
                        else
                        {
                            Hs = H.ToString();
                        }
                        if (M < 10)
                        {
                            Ms = "0" + M;
                        }
                        else
                        {
                            Ms = M.ToString();
                        }
                        if (S < 10)
                        {
                            Ss = "0" + S;
                        }
                        else
                        {
                            Ss = S.ToString();
                        }

                        string time = Hs + " : " + Ms + " : " + Ss;
                        //double Duration = EndTime.Subtract(StartTime).TotalMinutes;
                        //worksheet.Cells["I" + Row].Value = Math.Round(Duration, 2);
                        worksheet.Cells["I" + Row].Value = time;
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][5].ToString()) == false)
                        worksheet.Cells["J" + Row].Value = dataHolder.Rows[i][5].ToString();
                    else
                        worksheet.Cells["J" + Row].Value = "-";
                    Row++;
                    Sno++;
                }
            }

            ExcelRange r1, r2, r3;

            int noOfRows = 8 + dataHolder.Rows.Count + 2;
            worksheet.Cells["B8:B" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C8:C" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["D8:E" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["F8:H" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["I8:J" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            p.Save();

            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "BreakDownReport" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "BreakDownReport" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
        }

        //Idle Down Report
        public async void generateIdleReportExcel(string startDate, string EndtDate)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\IDLE_Report.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "IDLE_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "IDLE_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    ////TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                    IntoFile("Excel with same date is already open, please close it and try to generate!!!!");
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,StartDateTime,EndDateTime,MessageCodeID,CorrectedDate,EntryTime,Shift FROM [i_facility_tal].[dbo].tbllossofentry WHERE DoneWithRow = 1 AND CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' ORDER BY LossID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 8;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(y => y.MachineID == MachineID).FirstOrDefault();
                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineDispName;
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false)
                    {
                        DateTime startdate = Convert.ToDateTime(dataHolder.Rows[i][1]);
                        worksheet.Cells["D" + Row].Value = startdate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        DateTime Enddate = Convert.ToDateTime(dataHolder.Rows[i][2]);
                        worksheet.Cells["E" + Row].Value = Enddate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    //if (string.IsNullOrEmpty(dataHolder.Rows[i][5].ToString()) == false)
                    //{
                    //    DateTime EntryTime = Convert.ToDateTime(dataHolder.Rows[i][5]);
                    //    worksheet.Cells["F" + Row].Value = EntryTime.ToString("yyyy-MM-dd HH:mm:ss");
                    //}
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false && string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        DateTime StartTime = DateTime.Now;
                        StartTime = Convert.ToDateTime(dataHolder.Rows[i][1]);
                        DateTime EndTime = DateTime.Now;
                        EndTime = Convert.ToDateTime(dataHolder.Rows[i][2]);

                        TimeSpan ts = EndTime.Subtract(StartTime);
                        int H = ts.Hours;
                        int M = ts.Minutes;
                        int S = ts.Seconds;
                        string Hs = null, Ms = null, Ss = null;
                        if (H < 10)
                        {
                            Hs = "0" + H;
                        }
                        else
                        {
                            Hs = H.ToString();
                        }
                        if (M < 10)
                        {
                            Ms = "0" + M;
                        }
                        else
                        {
                            Ms = M.ToString();
                        }
                        if (S < 10)
                        {
                            Ss = "0" + S;
                        }
                        else
                        {
                            Ss = S.ToString();
                        }

                        string time = Hs + " : " + Ms + " : " + Ss;
                        //double Duration = EndTime.Subtract(StartTime).TotalMinutes;
                        //worksheet.Cells["I" + Row].Value = Math.Round(Duration, 2);
                        worksheet.Cells["I" + Row].Value = time;


                        //double Duration = EndTime.Subtract(StartTime).TotalMinutes;
                        //worksheet.Cells["I" + Row].Value = Math.Round(Duration, 2);
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][3].ToString()) == false)
                    {
                        int msgcd = Convert.ToInt32(dataHolder.Rows[i][3]);
                        var a = db.tbllossescodes.Where(m => m.LossCodeID == msgcd).SingleOrDefault();

                        if (a.LossCodesLevel == 1)
                        {
                            if (a.LossCode == "999")
                            {
                                worksheet.Cells["F" + Row].Value = a.MessageType;
                            }
                            else
                            {
                                worksheet.Cells["F" + Row].Value = a.LossCode;
                            }
                        }
                        else if (a.LossCodesLevel == 2)
                        {
                            int lossid = Convert.ToInt32(a.LossCodesLevel1ID);
                            var level1data = db.tbllossescodes.Where(m => m.LossCodeID == lossid).SingleOrDefault();
                            if (level1data.LossCode == "999")
                            {
                                worksheet.Cells["F" + Row].Value = level1data.MessageType;
                            }
                            else
                            {
                                worksheet.Cells["F" + Row].Value = level1data.LossCode;
                            }
                            worksheet.Cells["G" + Row].Value = a.LossCode;
                        }
                        else if (a.LossCodesLevel == 3)
                        {
                            int lossid2 = Convert.ToInt32(a.LossCodesLevel1ID);
                            int lossid3 = Convert.ToInt32(a.LossCodesLevel2ID);
                            var level1data = db.tbllossescodes.Where(m => m.LossCodeID == lossid2).SingleOrDefault();
                            var level2data = db.tbllossescodes.Where(m => m.LossCodeID == lossid3).SingleOrDefault();
                            if (level1data.LossCode == "999")
                            {
                                worksheet.Cells["F" + Row].Value = level1data.MessageType;
                            }
                            else
                            {
                                worksheet.Cells["F" + Row].Value = level1data.LossCode;
                            }
                            worksheet.Cells["G" + Row].Value = level2data.LossCode;
                            worksheet.Cells["H" + Row].Value = a.LossCode;
                        }
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][6].ToString()) == false)
                    {
                        worksheet.Cells["J" + Row].Value = dataHolder.Rows[i][6].ToString();
                    }
                    else
                        worksheet.Cells["J" + Row].Value = "-";
                    Row++;
                    Sno++;
                }
            }
            int noOfRows = 8 + dataHolder.Rows.Count + 2;
            worksheet.Cells["B8:B" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C8:C" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["D8:E" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["F8:H" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["I8:J" + noOfRows].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            ExcelRange r1, r2, r3;

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "IDLE_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "IDLE_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            ////return View();
        }

        //MUR_Report
        public async void generateMURReportExcel(string startDate, string EndtDate)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\MUR.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MUR" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "MUR" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");
            //var data = db.tblbreakdowns.Where(m => m.CorrectedDate >= FDate && m.CorrectedDate <= TDate);

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,MachineOnTime,OperatingTime FROM [i_facility_tal].[dbo].tblmimics WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' AND Shift = 'A' ORDER BY MachineID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();

            String sql2 = "SELECT MachineID,MachineOnTime,OperatingTime FROM [i_facility_tal].[dbo].tblmimics WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' AND Shift = 'B' ORDER BY MachineID ASC";
            SqlDataAdapter da2 = new SqlDataAdapter(sql2, mc.msqlConnection);
            DataTable dataHolder2 = new DataTable();
            da2.Fill(dataHolder2);
            mc.close();

            String sql3 = "SELECT MachineID,MachineOnTime,OperatingTime FROM [i_facility_tal].[dbo].tblmimics WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' AND Shift = 'C' ORDER BY MachineID ASC";
            SqlDataAdapter da3 = new SqlDataAdapter(sql3, mc.msqlConnection);
            DataTable dataHolder3 = new DataTable();
            da3.Fill(dataHolder3);
            mc.close();

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 9;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(m => m.MachineID == MachineID).FirstOrDefault();
                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineInvNo;
                    worksheet.Cells["D" + Row].Value = machineDetails.MachineMake;
                    worksheet.Cells["E" + Row].Value = machineDetails.MachineModel;

                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false)
                    {
                        double data = Convert.ToDouble(dataHolder.Rows[i][1]);
                        double poweron = (data / 480) * 100;
                        worksheet.Cells["F" + Row].Value = Math.Round(poweron, 2);
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        double data = Convert.ToDouble(dataHolder.Rows[i][2]);
                        double operatingtime = (data / 480) * 100;
                        worksheet.Cells["G" + Row].Value = Math.Round(operatingtime, 2);
                    }

                    if (i < dataHolder2.Rows.Count)
                    {
                        if (string.IsNullOrEmpty(dataHolder2.Rows[i][1].ToString()) == false)
                        {
                            double data = Convert.ToDouble(dataHolder2.Rows[i][1]);
                            double poweron = (data / 480) * 100;
                            worksheet.Cells["H" + Row].Value = Math.Round(poweron, 2);
                        }

                        if (string.IsNullOrEmpty(dataHolder2.Rows[i][2].ToString()) == false)
                        {
                            double data = Convert.ToDouble(dataHolder2.Rows[i][2]);
                            double operatingtime = (data / 480) * 100;
                            worksheet.Cells["I" + Row].Value = Math.Round(operatingtime, 2);
                        }
                    }
                    if (i < dataHolder3.Rows.Count)
                    {
                        if (string.IsNullOrEmpty(dataHolder3.Rows[i][1].ToString()) == false)
                        {
                            double data = Convert.ToDouble(dataHolder3.Rows[i][1]);
                            double poweron = (data / 480) * 100;
                            worksheet.Cells["J" + Row].Value = Math.Round(poweron, 2);
                        }
                        if (string.IsNullOrEmpty(dataHolder3.Rows[i][2].ToString()) == false)
                        {
                            double data = Convert.ToDouble(dataHolder3.Rows[i][2]);
                            double operatingtime = (data / 480) * 100;
                            worksheet.Cells["K" + Row].Value = Math.Round(operatingtime, 2);
                        }
                    }

                    Row++;
                    Sno++;
                }
            }

            ExcelRange r1, r2, r3;

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "MUR" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "MUR" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            //return View();
        }

        //Alarm_Report
        public async void generateAlarm_ReportExcel(string startDate, string EndtDate)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\Alarm_Report.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Alarm_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Alarm_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C5"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E5"].Value = toDate.ToString("dd-MM-yyyy");

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");
            //var data = db.tblbreakdowns.Where(m => m.CorrectedDate >= FDate && m.CorrectedDate <= TDate);

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,AlarmNumber,AlarmDesc,CreatedOn FROM [i_facility_tal].[dbo].tblpriorityalarms WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' ORDER BY AlarmID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 7;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(y => y.MachineID == MachineID).FirstOrDefault();
                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineInvNo;
                    worksheet.Cells["D" + Row].Value = machineDetails.MachineMake;
                    worksheet.Cells["E" + Row].Value = machineDetails.MachineModel;

                    if (string.IsNullOrEmpty(dataHolder.Rows[i][1].ToString()) == false)
                    {
                        worksheet.Cells["F" + Row].Value = dataHolder.Rows[i][1].ToString();
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        worksheet.Cells["G" + Row].Value = dataHolder.Rows[i][2].ToString();
                    }

                    if (string.IsNullOrEmpty(dataHolder.Rows[i][3].ToString()) == false)
                    {
                        try
                        {
                            DateTime date = Convert.ToDateTime(dataHolder.Rows[i][3]);
                            worksheet.Cells["H" + Row].Value = date.ToString("yyyy-MM-dd");
                        }
                        catch { }
                    }

                    Row++;
                    Sno++;
                }
            }

            ExcelRange r1, r2, r3;

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "Alarm_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "Alarm_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            //return View();
        }

        //HMIOperator
        public async void generateHMIOperatorReportExcel(string startDate, string EndtDate, int OperatorID)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\HMIOperator.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "HMIOperator" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "HMIOperator" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try
                {
                    worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");
            tbluser Operator = db.tblusers.Where(y => y.UserID == OperatorID).FirstOrDefault();
            worksheet.Cells["G6"].Value = Operator.DisplayName;

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");
            //var data = db.tblbreakdowns.Where(m => m.CorrectedDate >= FDate && m.CorrectedDate <= TDate);

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,OperatiorID,Shift,Date,Time,Project,PartNo,OperationNo,Rej_Qty,Work_Order_No,Target_Qty,Delivered_Qty,Prod_FAI FROM [i_facility_tal].[dbo].tblhmiscreen WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' AND OperatiorID=" + OperatorID + " ORDER BY HMIID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 8;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(y => y.MachineID == MachineID).FirstOrDefault();

                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineDispName;
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        worksheet.Cells["D" + Row].Value = dataHolder.Rows[i][2].ToString();
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][3].ToString()) == false)
                    {
                        DateTime date = Convert.ToDateTime(dataHolder.Rows[i][3]);
                        worksheet.Cells["E" + Row].Value = date.ToString("yyyy-MM-dd");
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][4].ToString()) == false)
                    {
                        string time = dataHolder.Rows[i][4].ToString();
                        DateTime date = Convert.ToDateTime(time);
                        worksheet.Cells["F" + Row].Value = date.ToString("HH:mm:ss");
                    }
                    //String sql1 = "SELECT MachineID0,OperatiorID1,Shift2,Date3,Time,Project,PartNo,OperationNo,Rej_Qty,Work_Order_No,Target_Qty,Delivered_Qty,Prod_FAI FROM tblhmiscreen WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' ORDER BY HMIID ASC";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][5].ToString()) == false)
                    {
                        worksheet.Cells["G" + Row].Value = dataHolder.Rows[i][5].ToString();
                    }
                    else
                        worksheet.Cells["G" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][6].ToString()) == false)
                    {
                        worksheet.Cells["H" + Row].Value = dataHolder.Rows[i][6].ToString();
                    }
                    else
                        worksheet.Cells["H" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][7].ToString()) == false)
                    {
                        worksheet.Cells["I" + Row].Value = dataHolder.Rows[i][7].ToString();
                    }
                    else
                        worksheet.Cells["I" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][9].ToString()) == false)
                    {
                        worksheet.Cells["J" + Row].Value = dataHolder.Rows[i][9].ToString();
                    }
                    else
                        worksheet.Cells["J" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][8].ToString()) == false)
                    {
                        worksheet.Cells["K" + Row].Value = dataHolder.Rows[i][8].ToString();
                    }
                    else
                        worksheet.Cells["K" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][10].ToString()) == false)
                    {
                        worksheet.Cells["L" + Row].Value = dataHolder.Rows[i][10].ToString();
                    }
                    else
                        worksheet.Cells["L" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][11].ToString()) == false)
                    {
                        worksheet.Cells["M" + Row].Value = dataHolder.Rows[i][11].ToString();
                    }
                    else
                        worksheet.Cells["M" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][12].ToString()) == false)
                    {
                        worksheet.Cells["N" + Row].Value = dataHolder.Rows[i][12].ToString();
                    }
                    else
                        worksheet.Cells["N" + Row].Value = "-";

                    Row++;
                    Sno++;
                }
            }

            ExcelRange r1, r2, r3;

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "HMIOperator" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "HMIOperator" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            //return View();
        }

        //HMIReport
        public async void generateHMIReportExcel(string startDate, string EndtDate)
        {
            DateTime frda = DateTime.Now;

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\HMIDate.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "HMIDate" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "HMIDate" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");
            //var data = db.tblbreakdowns.Where(m => m.CorrectedDate >= FDate && m.CorrectedDate <= TDate);

            MsqlConnection mc = new MsqlConnection();
            mc.open();
            String sql1 = "SELECT MachineID,OperatiorID,Shift,Date,Time,Project,PartNo,OperationNo,Rej_Qty,Work_Order_No,Target_Qty,Delivered_Qty,Prod_FAI FROM [i_facility_tal].[dbo].tblhmiscreen WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' ORDER BY HMIID ASC";
            SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection);
            DataTable dataHolder = new DataTable();
            da1.Fill(dataHolder);
            mc.close();

            if (dataHolder.Rows.Count != 0)
            {
                var Col = 'B';
                int Row = 8;
                int Sno = 1;
                for (int i = 0; i < dataHolder.Rows.Count; i++)
                {
                    int MachineID = Convert.ToInt32(dataHolder.Rows[i][0]);
                    int OperatorID = Convert.ToInt32(dataHolder.Rows[i][1]);
                    tblmachinedetail machineDetails = db.tblmachinedetails.Where(y => y.MachineID == MachineID).FirstOrDefault();

                    worksheet.Cells["B" + Row].Value = Sno;
                    worksheet.Cells["C" + Row].Value = machineDetails.MachineDispName;
                    tbluser Operator = db.tblusers.Where(y => y.UserID == OperatorID).FirstOrDefault();
                    worksheet.Cells["D" + Row].Value = Operator.DisplayName;
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][2].ToString()) == false)
                    {
                        worksheet.Cells["E" + Row].Value = dataHolder.Rows[i][2].ToString();
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][3].ToString()) == false)
                    {
                        DateTime date = Convert.ToDateTime(dataHolder.Rows[i][3]);
                        worksheet.Cells["F" + Row].Value = date.ToString("yyyy-MM-dd");
                    }
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][4].ToString()) == false)
                    {
                        string time = dataHolder.Rows[i][4].ToString();
                        DateTime date = Convert.ToDateTime(time);
                        worksheet.Cells["G" + Row].Value = date.ToString("HH:mm:ss");
                    }
                    //String sql1 = "SELECT MachineID0,OperatiorID1,Shift2,Date3,Time,Project,PartNo,OperationNo,Rej_Qty,Work_Order_No,Target_Qty,Delivered_Qty,Prod_FAI FROM tblhmiscreen WHERE CorrectedDate>='" + FDate + "' AND CorrectedDate<='" + TDate + "' ORDER BY HMIID ASC";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][5].ToString()) == false)
                    {
                        worksheet.Cells["H" + Row].Value = dataHolder.Rows[i][5].ToString();
                    }
                    else
                        worksheet.Cells["H" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][6].ToString()) == false)
                    {
                        worksheet.Cells["I" + Row].Value = dataHolder.Rows[i][6].ToString();
                    }
                    else
                        worksheet.Cells["I" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][7].ToString()) == false)
                    {
                        worksheet.Cells["J" + Row].Value = dataHolder.Rows[i][7].ToString();
                    }
                    else
                        worksheet.Cells["J" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][9].ToString()) == false)
                    {
                        worksheet.Cells["K" + Row].Value = dataHolder.Rows[i][9].ToString();
                    }
                    else
                        worksheet.Cells["K" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][8].ToString()) == false)
                    {
                        worksheet.Cells["L" + Row].Value = dataHolder.Rows[i][8].ToString();
                    }
                    else
                        worksheet.Cells["L" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][10].ToString()) == false)
                    {
                        worksheet.Cells["M" + Row].Value = dataHolder.Rows[i][10].ToString();
                    }
                    else
                        worksheet.Cells["M" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][11].ToString()) == false)
                    {
                        worksheet.Cells["N" + Row].Value = dataHolder.Rows[i][11].ToString();
                    }
                    else
                        worksheet.Cells["N" + Row].Value = "-";
                    if (string.IsNullOrEmpty(dataHolder.Rows[i][12].ToString()) == false)
                    {
                        worksheet.Cells["O" + Row].Value = dataHolder.Rows[i][12].ToString();
                    }
                    else
                        worksheet.Cells["O" + Row].Value = "-";

                    Row++;
                    Sno++;
                }
            }

            ExcelRange r1, r2, r3;

            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "HMIDate" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "HMIDate" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            //return View();
        }

        //UtilizationReport
        #region Utilization Report ShiftWise
        //public void UtilizationReportExcel(string startDate, string EndtDate, string Shift, string Shop, string WorkCenter)
        //{
        //    DateTime frda = DateTime.Now;
        //    if (string.IsNullOrEmpty(startDate) == true)
        //    {
        //        startDate = DateTime.Now.Date.ToString();
        //    }
        //    if (string.IsNullOrEmpty(EndtDate) == true)
        //    {
        //        EndtDate = startDate;
        //    }

        //    DateTime frmDate = Convert.ToDateTime(startDate);
        //    DateTime toDate = Convert.ToDateTime(EndtDate);

        //    double TotalDay = toDate.Subtract(frmDate).TotalDays;

        //    FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\Utilization_Report.xlsx");
        //    ExcelPackage templatep = new ExcelPackage(templateFile);
        //    ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

        //    String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
        //    //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

        //    bool exists = System.IO.Directory.Exists(FileDir);

        //    if (!exists)
        //        System.IO.Directory.CreateDirectory(FileDir);

        //    FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
        //    if (newFile.Exists)
        //    {
        //        try
        //        {
        //            newFile.Delete();  // ensures we create a new workbook
        //            newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
        //        }
        //        catch
        //        {
        //            //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
        //            ////return View();
        //        }
        //    }
        //    //Using the File for generation and populating it
        //    ExcelPackage p = null;
        //    p = new ExcelPackage(newFile);
        //    ExcelWorksheet worksheet = null;

        //    //Creating the WorkSheet for populating
        //    try
        //    {
        //        worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
        //    }
        //    catch { }

        //    if (worksheet == null)
        //    {
        //        worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
        //    }

        //    int sheetcount = p.Workbook.Worksheets.Count;
        //    p.Workbook.Worksheets.MoveToStart(sheetcount);

        //    worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //    worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //    worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
        //    worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

        //    DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);

        //    string FDate = frmDate.ToString("yyyy-MM-dd");
        //    string TDate = toDate.ToString("yyyy-MM-dd");

        //    {
        //        var Col = 'B';
        //        int Row = 8;
        //        int Sno = 1;
        //        for (int i = 0; i < TotalDay + 1; i++)
        //        {
        //            if (Shift == "No Use")
        //            {
        //                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");
        //                DataTable machin = new DataTable();
        //                if (Shop == "No Use" && WorkCenter == "No Use")
        //                {
        //                    MsqlConnection mc = new MsqlConnection();
        //                    mc.open();
        //                    String query1 = "SELECT distinct MachineID From tbldailyprodstatus WHERE CorrectedDate='" + dateforMachine + "' AND IsDeleted=0";
        //                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);

        //                    da1.Fill(machin);
        //                    mc.close();
        //                }
        //                else
        //                {
        //                    MsqlConnection mc = new MsqlConnection();
        //                    mc.open();
        //                    String query1 = "SELECT MachineID From tblmachinedetails WHERE MachineDispName='" + WorkCenter + "' AND IsDeleted=0  AND ShopNo='" + Shop + "'";
        //                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);

        //                    da1.Fill(machin);
        //                    mc.close();
        //                }
        //                for (int n = 0; n < machin.Rows.Count; n++)
        //                {
        //                    var shft = 'A';
        //                    worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

        //                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
        //                    tblmachinedetail machineDetails = db.tblmachinedetails.Find(MachineID);

        //                    for (int j = 0; j < 3; j++)
        //                    {
        //                        worksheet.Cells["B" + Row].Value = Sno;

        //                        worksheet.Cells["D" + Row].Value = shft.ToString();
        //                        double green, red, yellow, blue, setup = 0;
        //                        if (shft.ToString() == "C")
        //                        {
        //                            blue = GetOPIDleBreakDownForShift3(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                            yellow = GetOPIDleBreakDownForShift3(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                            setup = GetOPIDleBreakDownSetupForShift3(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                            green = GetOPIDleBreakDownForShift3(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
        //                            red = GetOPIDleBreakDownForShift3(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
        //                        }
        //                        else
        //                        {
        //                            blue = GetOPIDleBreakDown(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                            yellow = GetOPIDleBreakDown(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                            setup = GetOPIDleBreakDownSetup(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                            green = GetOPIDleBreakDown(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
        //                            red = GetOPIDleBreakDown(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
        //                        }
        //                        double TotalMinutes = green + setup + (yellow - setup) + red + blue;
        //                        double Diff = 480 - TotalMinutes;
        //                        worksheet.Cells["G" + Row].Value = green;
        //                        worksheet.Cells["H" + Row].Value = setup;
        //                        worksheet.Cells["I" + Row].Value = yellow - setup;
        //                        worksheet.Cells["J" + Row].Value = red;
        //                        worksheet.Cells["K" + Row].Value = blue + Diff;
        //                        worksheet.Cells["E" + Row].Value = machineDetails.ShopNo.ToString();
        //                        worksheet.Cells["F" + Row].Value = machineDetails.MachineDispName.ToString();

        //                        //Utilisation
        //                        worksheet.Cells["L" + Row].Value = Math.Round((green / 480) * 100, 2);
        //                        ////Availability
        //                        //double basevalue = 480 - (blue + Diff);
        //                        //if (basevalue == 0)
        //                        //    basevalue = 1;
        //                        ////if (red != 0)
        //                        ////{
        //                        //double val = Convert.ToDouble(Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2));
        //                        //if (val > 0)
        //                        //    worksheet.Cells["M" + Row].Value = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);
        //                        //else
        //                        //    worksheet.Cells["M" + Row].Value = 0;
        //                        ////}
        //                        ////else
        //                        //worksheet.Cells["M" + Row].Value = 0;
        //                        //double avalibality = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);

        //                        //Availability
        //                        double basevalue = 480;
        //                        if (basevalue == 0)
        //                            basevalue = 1;
        //                        //if (red != 0)
        //                        //{
        //                        double val = Convert.ToDouble(Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2));
        //                        if (val > 0)
        //                            worksheet.Cells["M" + Row].Value = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);
        //                        else
        //                            worksheet.Cells["M" + Row].Value = 0;
        //                        //}
        //                        //else
        //                        //worksheet.Cells["M" + Row].Value = 0;
        //                        double avalibality = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);


        //                        //Efficiency
        //                        int val1 = 0;
        //                        double Zero = 480 - (blue + Diff);
        //                        if (Zero != 0)
        //                            val1 = Convert.ToInt32(Math.Round((green) / (480 - (blue + Diff)) * 100, 2));
        //                        if (val1 > 0)
        //                            worksheet.Cells["N" + Row].Value = Math.Round((green) / (480 - (blue + Diff)) * 100, 2);
        //                        else
        //                            worksheet.Cells["N" + Row].Value = 0;
        //                        double Efficiency = Math.Round((green) / (480 - (blue + Diff)) * 100, 2);

        //                        //Quality
        //                        double qty = GetQuality(shft.ToString(), UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        qty = Math.Round((qty * 100), 2);
        //                        worksheet.Cells["O" + Row].Value = qty;

        //                        //OEE
        //                        //worksheet.Cells["P" + Row].Value = Math.Round((avalibality / 100 * Efficiency / 100 * qty / 100) * 100, 2);
        //                        worksheet.Cells["P" + Row].Value = Math.Round((((avalibality / 100) * (Efficiency / 100) * (qty / 100))) * 100, 2);
        //                        shft++;
        //                        Row++;
        //                        Sno++;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");
        //                DataTable machin = new DataTable();
        //                if (Shop == "No Use" && WorkCenter == "No Use")
        //                {
        //                    MsqlConnection mc = new MsqlConnection();
        //                    mc.open();
        //                    String query1 = "SELECT distinct MachineID From tbldailyprodstatus WHERE CorrectedDate='" + dateforMachine + "' AND IsDeleted=0";
        //                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //                    da1.Fill(machin);
        //                    mc.close();
        //                }
        //                else
        //                {
        //                    MsqlConnection mc = new MsqlConnection();
        //                    mc.open();
        //                    String query1 = "SELECT MachineID From tblmachinedetails WHERE MachineDispName='" + WorkCenter + "' AND IsDeleted=0  AND ShopNo='" + Shop + "'";
        //                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);

        //                    da1.Fill(machin);
        //                    mc.close();
        //                }
        //                for (int n = 0; n < machin.Rows.Count; n++)
        //                {

        //                    int MachineID = Convert.ToInt32(machin.Rows[n][0]);
        //                    tblmachinedetail machineDetails = db.tblmachinedetails.Find(MachineID);
        //                    worksheet.Cells["B" + Row].Value = Sno;
        //                    worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");
        //                    worksheet.Cells["D" + Row].Value = Shift;
        //                    double green, red, yellow, blue, setup = 0;
        //                    if (Shift == "C")
        //                    {
        //                        blue = GetOPIDleBreakDownForShift3(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                        yellow = GetOPIDleBreakDownForShift3(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                        setup = GetOPIDleBreakDownSetupForShift3(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        green = GetOPIDleBreakDownForShift3(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
        //                        red = GetOPIDleBreakDownForShift3(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
        //                    }
        //                    else
        //                    {
        //                        blue = GetOPIDleBreakDown(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");
        //                        yellow = GetOPIDleBreakDown(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
        //                        setup = GetOPIDleBreakDownSetup(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                        green = GetOPIDleBreakDown(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
        //                        red = GetOPIDleBreakDown(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
        //                    }
        //                    double TotalMinutes = green + setup + (yellow - setup) + red + blue;
        //                    double Diff = 480 - TotalMinutes;
        //                    worksheet.Cells["G" + Row].Value = green;
        //                    //Setting
        //                    worksheet.Cells["H" + Row].Value = setup;
        //                    //Idle
        //                    worksheet.Cells["I" + Row].Value = yellow - setup;
        //                    worksheet.Cells["J" + Row].Value = red;
        //                    worksheet.Cells["K" + Row].Value = blue + Diff;
        //                    worksheet.Cells["E" + Row].Value = machineDetails.ShopNo.ToString();
        //                    worksheet.Cells["F" + Row].Value = machineDetails.MachineDispName.ToString();

        //                    //Utilisation
        //                    worksheet.Cells["L" + Row].Value = Math.Round((green / 480) * 100, 2);

        //                    //Availability
        //                    double basevalue = 480 - (blue + Diff);
        //                    if (basevalue == 0)
        //                        basevalue = 1;
        //                    //if (red != 0)
        //                    worksheet.Cells["M" + Row].Value = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);
        //                    //else
        //                    //worksheet.Cells["M" + Row].Value = 0;
        //                    double avalibality = Math.Round(((480 - red - (blue + Diff)) / (basevalue)) * 100, 2);

        //                    //Efficiency
        //                    worksheet.Cells["N" + Row].Value = Math.Round((green) / (480 - (blue + Diff)) * 100, 2);
        //                    double Efficiency = Math.Round((green) / (480 - (blue + Diff)) * 100, 2);

        //                    //Quality
        //                    double qty = GetQuality(Shift, UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
        //                    qty = Math.Round((qty * 100), 2);
        //                    worksheet.Cells["O" + Row].Value = qty;

        //                    //OEE
        //                    worksheet.Cells["P" + Row].Value = Math.Round((((avalibality / 100) * (Efficiency / 100) * (qty / 100))) * 100, 2);

        //                    Row++;
        //                    Sno++;
        //                }
        //            }
        //            UsedDateForExcel = UsedDateForExcel.AddDays(+1);
        //        }
        //    }

        //    ExcelRange r1, r2, r3;

        //    p.Save();
        //    //Downloding Excel
        //    string path1 = System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx");
        //    System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
        //    string Outgoingfile = "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx";
        //    if (file1.Exists)
        //    {
        //        //Response.Clear();
        //        //Response.ClearContent();
        //        //Response.ClearHeaders();
        //        //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
        //        //Response.AddHeader("Content-Length", file1.Length.ToString());
        //        //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        //Response.WriteFile(file1.FullName);
        //        //Response.Flush();
        //        //Response.Close();
        //    }
        //    //return View();
        //}

        //public int GetOPIDleBreakDownForShift3(string Shift, string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");
        //    string StartTime = datetime;
        //    string EndTime = datetime;

        //    string EndTime1 = datetime;
        //    string StartTime1 = datetime;
        //    //int CurrentHour = DateTime.Now.Hour;
        //    //if (CurrentHour < 6)
        //    {
        //        datetime = currentdate.ToString("yyyy-MM-dd");
        //        StartTime = datetime + " 22:00:00";
        //        EndTime = datetime + " 23:59:59";
        //    }
        //    int[] count = new int[4];
        //    MsqlConnection mc = new MsqlConnection();
        //    //operating
        //    DataTable OP = new DataTable();
        //    //if (CurrentHour < 6)
        //    {
        //        mc.open();
        //        String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTime + "' AND EndTime <='" + EndTime + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //        SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //        da1.Fill(OP);
        //        mc.close();

        //        datetime = currentdate.AddDays(+1).ToString("yyyy-MM-dd");
        //        StartTime1 = datetime + " 00:00:00";
        //        EndTime1 = datetime + " 05:59:59";
        //        mc.open();
        //        String query11 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTime1 + "' AND EndTime <='" + EndTime1 + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //        SqlDataAdapter da11 = new SqlDataAdapter(query11, mc.msqlConnection);
        //        DataTable OP1 = new DataTable();
        //        da11.Fill(OP1);
        //        mc.close();

        //        OP.Rows[0][0] = Convert.ToInt32(OP.Rows[0][0]) + Convert.ToInt32(OP1.Rows[0][0]);
        //    }
        //    if (OP.Rows.Count != 0)
        //    {
        //        count[0] = Convert.ToInt32(OP.Rows[0][0]);
        //    }
        //    //return count[0];
        //}

        //public int GetOPIDleBreakDown(string Shift, string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");
        //    string StartTime = datetime;
        //    string EndTime = datetime;
        //    if (Shift == "A")
        //    {
        //        StartTime = StartTime + " 06:00:00";
        //        EndTime = EndTime + " 14:00:00";
        //    }
        //    if (Shift == "B")
        //    {
        //        StartTime = StartTime + " 14:00:00";
        //        EndTime = EndTime + " 22:00:00";
        //    }
        //    int[] count = new int[4];
        //    MsqlConnection mc = new MsqlConnection();
        //    mc.open();
        //    //operating
        //    mc.open();
        //    String query1 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTime + "' AND EndTime <='" + EndTime + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //    DataTable OP = new DataTable();
        //    da1.Fill(OP);
        //    mc.close();
        //    if (OP.Rows.Count != 0)
        //    {
        //        count[0] = Convert.ToInt32(OP.Rows[0][0]);
        //    }
        //    //return count[0];
        //}

        //public double GetQuality(string Shift, string CorrectedDate, int MachineID)
        //{
        //    double[] count = new double[4];
        //    MsqlConnection mc = new MsqlConnection();
        //    mc.open();
        //    //operating
        //    mc.open();
        //    String query1 = "SELECT Delivered_Qty,Rej_Qty From tblhmiscreen WHERE Shift='" + Shift + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " ";
        //    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //    DataTable OP = new DataTable();
        //    da1.Fill(OP);
        //    mc.close();
        //    if (OP.Rows.Count != 0)
        //    {
        //        double Sumdelvrd = 0;
        //        double Sumrej = 0;
        //        for (int i = 0; i < OP.Rows.Count; i++)
        //        {
        //            double delivdqty = 0;
        //            if (string.IsNullOrEmpty(OP.Rows[i][0].ToString()) == false)
        //                delivdqty = Convert.ToDouble(OP.Rows[i][0]);
        //            double Rejqty = 0;
        //            if (string.IsNullOrEmpty(OP.Rows[i][1].ToString()) == false)
        //                Rejqty = Convert.ToDouble(OP.Rows[i][1]);
        //            Sumdelvrd = Sumdelvrd + delivdqty;
        //            Sumrej = Sumrej + Rejqty;
        //        }
        //        if (Sumdelvrd != 0)
        //            count[0] = ((Sumdelvrd) / (Sumdelvrd + Sumrej));
        //        else
        //            count[0] = ((Sumdelvrd - Sumrej) / 1);
        //    }
        //    //return count[0];
        //}

        //public int GetOPIDleBreakDownSetup(string Shift, string CorrectedDate, int MachineID)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");
        //    string StartTime = datetime;
        //    string EndTime = datetime;
        //    if (Shift == "A")
        //    {
        //        StartTime = StartTime + " 06:00:00";
        //        EndTime = EndTime + " 14:00:00";
        //    }
        //    if (Shift == "B")
        //    {
        //        StartTime = StartTime + " 14:00:00";
        //        EndTime = EndTime + " 22:00:00";
        //    }
        //    int[] count = new int[4];
        //    count[0] = 0;
        //    MsqlConnection mc = new MsqlConnection();
        //    mc.open();
        //    //operating
        //    mc.open();
        //    String query1 = "SELECT StartDateTime,EndDateTime From tbllossofentry WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND StartDateTime >='" + StartTime + "' AND EndDateTime <='" + EndTime + "'";
        //    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //    DataTable OP = new DataTable();
        //    da1.Fill(OP);
        //    mc.close();
        //    string StartTimedt = null;
        //    string EndTimedt = null;
        //    if (OP.Rows.Count != 0)
        //    {
        //        for (int i = 0; i < OP.Rows.Count; i++)
        //        {
        //            DateTime stdt = DateTime.Now;
        //            DateTime enddt = DateTime.Now;
        //            if (string.IsNullOrEmpty(OP.Rows[i][0].ToString()) == false)
        //            {
        //                StartTimedt = OP.Rows[i][0].ToString();
        //                stdt = Convert.ToDateTime(StartTimedt);
        //                StartTimedt = stdt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }
        //            if (string.IsNullOrEmpty(OP.Rows[i][1].ToString()) == false)
        //            {
        //                EndTimedt = OP.Rows[i][1].ToString();
        //                enddt = Convert.ToDateTime(EndTimedt);
        //                EndTimedt = enddt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }
        //            if (StartTimedt != null && EndTimedt != null)
        //            {
        //                mc.open();
        //                String query2 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTimedt + "' AND EndTime <='" + EndTimedt + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + "";
        //                SqlDataAdapter da3 = new SqlDataAdapter(query2, mc.msqlConnection);
        //                DataTable OP1 = new DataTable();
        //                da3.Fill(OP1);
        //                mc.close();
        //                if (string.IsNullOrEmpty(OP1.Rows[0][0].ToString()) == false)
        //                    count[0] = count[0] + Convert.ToInt32(OP1.Rows[0][0]);
        //            }
        //        }
        //    }
        //    //return count[0];
        //}

        //public int GetOPIDleBreakDownSetupForShift3(string Shift, string CorrectedDate, int MachineID)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");
        //    string StartTime = datetime;
        //    string EndTime = datetime;

        //    string EndTime1 = datetime;
        //    string StartTime1 = datetime;
        //    //int CurrentHour = DateTime.Now.Hour;
        //    //if (CurrentHour < 6)
        //    {
        //        datetime = currentdate.ToString("yyyy-MM-dd");
        //        StartTime = datetime + " 22:00:00";
        //        EndTime = datetime + " 23:59:59";
        //    }
        //    int[] count = new int[4];
        //    MsqlConnection mc = new MsqlConnection();
        //    //operating
        //    DataTable OP = new DataTable();
        //    DataTable OP1 = new DataTable();
        //    //if (CurrentHour < 6)
        //    {
        //        mc.open();
        //        String query1 = "SELECT StartDateTime,EndDateTime From tbllossofentry WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND StartDateTime >='" + StartTime + "' AND EndDateTime <='" + EndTime + "'";
        //        SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
        //        da1.Fill(OP);
        //        mc.close();

        //        datetime = currentdate.AddDays(+1).ToString("yyyy-MM-dd");
        //        StartTime1 = datetime + " 00:00:00";
        //        EndTime1 = datetime + " 05:59:59";
        //        mc.open();
        //        String query11 = "SELECT StartDateTime,EndDateTime From tbllossofentry WHERE StartDateTime >='" + StartTime1 + "' AND EndDateTime <='" + EndTime1 + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + "";
        //        SqlDataAdapter da11 = new SqlDataAdapter(query11, mc.msqlConnection);

        //        da11.Fill(OP1);
        //        mc.close();

        //        //OP.Rows[0][0] = Convert.ToInt32(OP.Rows[0][0]) + Convert.ToInt32(OP1.Rows[0][0]);
        //    }
        //    string StartTimedt = null;
        //    string EndTimedt = null;
        //    if (OP.Rows.Count != 0)
        //    {
        //        for (int i = 0; i < OP.Rows.Count; i++)
        //        {
        //            DateTime stdt = DateTime.Now;
        //            DateTime enddt = DateTime.Now;
        //            if (string.IsNullOrEmpty(OP.Rows[i][0].ToString()) == false)
        //            {
        //                StartTimedt = OP.Rows[i][0].ToString();
        //                stdt = Convert.ToDateTime(StartTimedt);
        //                StartTimedt = stdt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }
        //            if (string.IsNullOrEmpty(OP.Rows[i][1].ToString()) == false)
        //            {
        //                EndTimedt = OP.Rows[i][1].ToString();
        //                enddt = Convert.ToDateTime(EndTimedt);
        //                EndTimedt = enddt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }

        //            mc.open();
        //            String query2 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTimedt + "' AND EndTime <='" + EndTimedt + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='yellow'";
        //            SqlDataAdapter da3 = new SqlDataAdapter(query2, mc.msqlConnection);
        //            DataTable OP3 = new DataTable();
        //            da3.Fill(OP3);
        //            mc.close();
        //            if (string.IsNullOrEmpty(OP3.Rows[i][0].ToString()) == false)
        //                count[0] = count[0] + Convert.ToInt32(OP3.Rows[i][0]);
        //        }
        //        for (int i = 0; i < OP1.Rows.Count; i++)
        //        {
        //            DateTime stdt = DateTime.Now;
        //            DateTime enddt = DateTime.Now;

        //            if (string.IsNullOrEmpty(OP1.Rows[i][0].ToString()) == false)
        //            {
        //                StartTimedt = OP1.Rows[i][0].ToString();
        //                stdt = Convert.ToDateTime(StartTimedt);
        //                StartTimedt = stdt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }
        //            if (string.IsNullOrEmpty(OP1.Rows[i][1].ToString()) == false)
        //            {
        //                EndTimedt = OP1.Rows[i][1].ToString();
        //                enddt = Convert.ToDateTime(EndTimedt);
        //                EndTimedt = enddt.ToString("yyyy-MM-dd HH:mm:ss");
        //            }

        //            if (StartTimedt != null && EndTimedt != null)
        //            {
        //                mc.open();
        //                String query2 = "SELECT count(ID) From tbldailyprodstatus WHERE StartTime >='" + StartTimedt + "' AND EndTime <='" + EndTimedt + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + "";
        //                SqlDataAdapter da3 = new SqlDataAdapter(query2, mc.msqlConnection);
        //                DataTable OP4 = new DataTable();
        //                da3.Fill(OP4);
        //                mc.close();
        //                if (string.IsNullOrEmpty(OP4.Rows[0][0].ToString()) == false)
        //                    count[0] = count[0] + Convert.ToInt32(OP4.Rows[0][0]);
        //            }
        //        }
        //    }
        //    //return count[0];
        //}
        #endregion

        //UtilizationReport DayWise
        public async void UtilizationReportExcel(string startDate, string EndtDate, string Shop, string WorkCenter, string TimeType)
        {
            string lowestLevelLoss = null;
            DateTime frda = DateTime.Now;
            if (string.IsNullOrEmpty(startDate) == true)
            {
                startDate = DateTime.Now.Date.ToString();
            }
            if (string.IsNullOrEmpty(EndtDate) == true)
            {
                EndtDate = startDate;
            }

            DateTime frmDate = Convert.ToDateTime(startDate);
            DateTime toDate = Convert.ToDateTime(EndtDate);

            double TotalDay = toDate.Subtract(frmDate).TotalDays;

            FileInfo templateFile = new FileInfo(@"C:\TataReport\Templet\Utilization_Report.xlsx");
            ExcelPackage templatep = new ExcelPackage(templateFile);
            ExcelWorksheet Templatews = templatep.Workbook.Worksheets[1];

            String FileDir = @"C:\TataReport\AutoReportFiles\" + System.DateTime.Now.ToString("yyyy-MM-dd");
            //String FileDir = @"C:\inetpub\ContiAndonWebApp\Reports\" + System.DateTime.Now.ToString("yyyy");

            bool exists = System.IO.Directory.Exists(FileDir);

            if (!exists)
                System.IO.Directory.CreateDirectory(FileDir);

            FileInfo newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //+ " to " + toda.ToString("yyyy-MM-dd") 
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx")); //" to " + toda.ToString("yyyy-MM-dd") + 
                }
                catch
                {
                    //TempData["Excelopen"] = "Excel with same date is already open, please close it and try to generate!!!!";
                    ////return View();
                }
            }
            //Using the File for generation and populating it
            ExcelPackage p = null;
            p = new ExcelPackage(newFile);
            ExcelWorksheet worksheet = null;

            //Creating the WorkSheet for populating
            try
            {
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
            }
            catch { }

            if (worksheet == null)
            {
                try{
                worksheet = p.Workbook.Worksheets.Add(System.DateTime.Now.ToString("dd-MM-yyyy"), Templatews);
                }
                catch (Exception e)
                { }
            }

            int sheetcount = p.Workbook.Worksheets.Count;
            p.Workbook.Worksheets.MoveToStart(sheetcount);

            worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells["C6"].Value = frmDate.ToString("dd-MM-yyyy");
            worksheet.Cells["E6"].Value = toDate.ToString("dd-MM-yyyy");

            DateTime UsedDateForExcel = Convert.ToDateTime(frmDate);

            string FDate = frmDate.ToString("yyyy-MM-dd");
            string TDate = toDate.ToString("yyyy-MM-dd");

            {
                var Col = 'B';
                int Row = 8;
                int Sno = 1;
                for (int i = 0; i < TotalDay + 1; i++)
                {

                    string dateforMachine = UsedDateForExcel.ToString("yyyy-MM-dd");
                    DataTable machin = new DataTable();

                    MsqlConnection mc = new MsqlConnection();
                    mc.open();
                    String query1 = null;
                    if (Shop == "No Use" && WorkCenter == "No Use")
                    {
                        query1 = "SELECT distinct MachineID From [i_facility_tal].[dbo].tbldailyprodstatus WHERE CorrectedDate = '" + dateforMachine + "' AND IsDeleted=0";
                    }
                    else if (Shop != "No Use" && WorkCenter != "No Use")
                    {
                        query1 = "SELECT MachineID From [i_facility_tal].[dbo].tblmachinedetails WHERE MachineDispName='" + WorkCenter + "' AND IsDeleted=0  AND ShopNo = '" + Shop + "'";
                    }
                    else if (Shop != "No Use")
                    {
                        query1 = "SELECT MachineID From [i_facility_tal].[dbo].tblmachinedetails WHERE IsDeleted=0  AND ShopNo='" + Shop + "'";
                    }
                    else if (WorkCenter != "No Use")
                    {
                        query1 = "SELECT MachineID From [i_facility_tal].[dbo].tblmachinedetails WHERE IsDeleted=0  AND MachineDispName='" + WorkCenter + "'";
                    }
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(machin);
                    mc.close();

                    for (int n = 0; n < machin.Rows.Count; n++)
                    {
                        var shft = 'A';
                        worksheet.Cells["C" + Row].Value = UsedDateForExcel.ToString("yyyy-MM-dd");

                        int MachineID = Convert.ToInt32(machin.Rows[n][0]);
                        tblmachinedetail machineDetails = db.tblmachinedetails.Where(y => y.MachineID == MachineID).FirstOrDefault();

                        worksheet.Cells["B" + Row].Value = Sno;
                        double AvaillabilityFactor, EfficiencyFactor, QualityFactor;
                        double green, red, yellow, blue = 0, MinorLoss = 0, setup = 0, scrap = 0, NOP = 0, OperatingTime = 0, DownTimeBreakdown = 0, ROALossess = 0, AvailableTime = 0, SettingTime = 0, PlannedDownTime = 0, UnPlannedDownTime = 0;
                        double SummationOfSCTvsPP = 0, MinorLosses = 0, ROPLosses = 0;
                        double ScrapQtyTime = 0, ReWOTime = 0, ROQLosses = 0;
                        //GH PH NBH
                        if (TimeType == "GH")
                        {
                            AvailableTime = 24 * 60; //24Hours to Minutes
                        }

                        //
                        //yellow = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
                        //MinorLoss = GetMinorLoss(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
                        //setup = GetOPIDleBreakDownSetup(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        //green = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "green");
                        //red = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "red");
                        //NOP = green - MinorLoss;
                        //scrap = GetScrap(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, NOP);

                        //######
                        //MinorLosses = GetMinorLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "yellow");
                        //blue = GetOPIDleBreakDown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "blue");

                        ////Availability
                        //SettingTime = GetSettingTime(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        //ROALossess = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROA");
                        //DownTimeBreakdown = GetDownTimeBreakdown(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);


                        ////Performance
                        //SummationOfSCTvsPP = GetSummationOfSCTvsPP(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        ////ROPLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROP");

                        ////Quality
                        //ScrapQtyTime = GetScrapQtyTimeOfWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        //ReWOTime = GetScrapQtyTimeOfRWO(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        ////ROQLosses = GetDownTimeLosses(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID, "ROQ");
                        //######

                        //New Logic . Take values from 
                        DateTime DateTimeValue = Convert.ToDateTime(UsedDateForExcel.ToString("yyyy-MM-dd") + " " + "00:00:00");
                        var OEEData = db.tbloeedashboardvariables.Where(m => m.StartDate == DateTimeValue && m.WCID == MachineID).FirstOrDefault();
                        if (OEEData != null)
                        {
                            MinorLosses = Convert.ToDouble(OEEData.MinorLosses);
                            blue = Convert.ToDouble(OEEData.Blue);
                            SettingTime = Convert.ToDouble(OEEData.SettingTime);
                            ROALossess = Convert.ToDouble(OEEData.ROALossess);
                            DownTimeBreakdown = Convert.ToDouble(OEEData.DownTimeBreakdown);
                            SummationOfSCTvsPP = Convert.ToDouble(OEEData.SummationOfSCTvsPP);

                            ScrapQtyTime = Convert.ToDouble(OEEData.ScrapQtyTime);
                            ReWOTime = Convert.ToDouble(OEEData.ReWOTime);
                        }
                        else
                        {

                        }

                        // OperatingTime = AvailableTime - (ROALossess + DownTimeBreakdown + blue + MinorLosses + ROPLosses + ROQLosses);
                        OperatingTime = AvailableTime - (ROALossess + DownTimeBreakdown + blue + MinorLosses);

                        setup = SettingTime;
                        //double TotalMinutes = green + setup + (yellow - setup) + red + blue;
                        double TotalMinutes = OperatingTime + SettingTime + (ROALossess - SettingTime) + MinorLosses + DownTimeBreakdown + blue;
                        //double Diff = 1440 - TotalMinutes;

                        //for 3 shifts so 
                        double Diff = (8 * 3 * 60) - TotalMinutes;

                        worksheet.Cells["F" + Row].Value = Math.Round(OperatingTime, 2);
                        worksheet.Cells["G" + Row].Value = Math.Round(SettingTime, 2);
                        worksheet.Cells["H" + Row].Value = Math.Round((ROALossess - SettingTime), 2);
                        worksheet.Cells["I" + Row].Value = Math.Round(MinorLosses, 2);
                        //worksheet.Cells["J" + Row].Value = Math.Round(ROALossess, 2);
                        //worksheet.Cells["K" + Row].Value = Math.Round(ROPLosses, 2);
                        //worksheet.Cells["L" + Row].Value = Math.Round(ROQLosses, 2);
                        worksheet.Cells["J" + Row].Value = DownTimeBreakdown;
                        worksheet.Cells["K" + Row].Value = blue; //Dont try to adjust now " + Diff ".
                        worksheet.Cells["D" + Row].Value = machineDetails.ShopNo.ToString();
                        worksheet.Cells["E" + Row].Value = machineDetails.MachineDispName.ToString();

                        #region commented
                        ////Utilisation
                        //worksheet.Cells["K" + Row].Value = Math.Round((green / 1440) * 100, 2);


                        ////Availability
                        //double basevalue = 1440;
                        //if (basevalue == 0)
                        //    basevalue = 1;
                        ////if (red != 0)
                        ////{
                        //double val = Convert.ToDouble(Math.Round(((1440 - red - (blue + Diff)) / (basevalue)) * 100, 2));
                        //if (val > 0)
                        //    worksheet.Cells["L" + Row].Value = Math.Round(((1440 - red - (blue + Diff)) / (basevalue)) * 100, 2);
                        //else
                        //    worksheet.Cells["L" + Row].Value = 0;
                        ////}
                        ////else
                        ////worksheet.Cells["M" + Row].Value = 0;
                        //double avalibality = Math.Round(((1440 - red - (blue + Diff)) / (basevalue)) * 100, 2);

                        ////Efficiency
                        //int val1 = 0;
                        //double Zero = 1440 - (blue + Diff);
                        //if (Zero != 0)
                        //    val1 = Convert.ToInt32(Math.Round((green) / (1440 - (blue + Diff)) * 100, 2));
                        //if (val1 > 0)
                        //    worksheet.Cells["M" + Row].Value = Math.Round((green) / (1440 - (blue + Diff)) * 100, 2);
                        //else
                        //    worksheet.Cells["M" + Row].Value = 0;
                        //double Efficiency = Math.Round((green) / (1440 - (blue + Diff)) * 100, 2);

                        ////Quality
                        //double qty = GetQuality(UsedDateForExcel.ToString("yyyy-MM-dd"), MachineID);
                        //qty = Math.Round((qty * 100), 2);
                        //worksheet.Cells["N" + Row].Value = qty;
                        #endregion

                        #region 2016-08-04 New Formulas.

                        //Utilisation // Whole day duration is 24*60 = 1440 minutes
                        double val = OperatingTime / AvailableTime;
                        if (val > 0)
                            worksheet.Cells["L" + Row].Value = Math.Round(val * 100, 2);
                        else
                            worksheet.Cells["L" + Row].Value = 0;

                        //Availablity
                        if (val > 0)
                            worksheet.Cells["M" + Row].Value = Math.Round(val * 100, 2);
                        else
                            worksheet.Cells["M" + Row].Value = 0;
                        AvaillabilityFactor = Math.Round(val * 100, 2);




                        //Performance
                        if (SummationOfSCTvsPP == -1 || SummationOfSCTvsPP == 0)
                        {
                            EfficiencyFactor = 100;
                            worksheet.Cells["N" + Row].Value = Math.Round(EfficiencyFactor, 2);
                        }
                        else
                        {
                            EfficiencyFactor = Math.Round((SummationOfSCTvsPP / (OperatingTime)) * 100, 2);
                            if (EfficiencyFactor >= 0 && EfficiencyFactor <= 100)
                                worksheet.Cells["N" + Row].Value = EfficiencyFactor;
                            else if (EfficiencyFactor > 100)
                            {
                                EfficiencyFactor = 100;
                                worksheet.Cells["N" + Row].Value = 100;
                            }
                            else if (EfficiencyFactor < 0)
                            {
                                EfficiencyFactor = 0;
                                worksheet.Cells["N" + Row].Value = 0;
                            }
                        }

                        //Quality
                        QualityFactor = (OperatingTime - ScrapQtyTime - ReWOTime) / OperatingTime;
                        if (QualityFactor >= 0 && QualityFactor <= 100)
                        {
                            worksheet.Cells["O" + Row].Value = Math.Round(QualityFactor * 100, 2);
                        }
                        else if (QualityFactor > 100)
                        {
                            QualityFactor = 100;
                            worksheet.Cells["O" + Row].Value = 100;
                        }
                        else if (QualityFactor < 0)
                        {
                            QualityFactor = 0;
                            worksheet.Cells["O" + Row].Value = 0;
                        }


                        //OEE
                        if (AvaillabilityFactor <= 0 || EfficiencyFactor <= 0 || QualityFactor <= 0)
                        {
                            worksheet.Cells["P" + Row].Value = 0;
                        }
                        else
                        {
                            val = Math.Round((AvaillabilityFactor / 100) * (EfficiencyFactor / 100) * (QualityFactor / 100) * 100, 2);
                            if (val >= 0 && val <= 100)
                            {
                                worksheet.Cells["P" + Row].Value = Math.Round(val * 100, 2);
                            }
                            else if (val > 100)
                            {
                                worksheet.Cells["P" + Row].Value = 100;
                            }
                            else if (val < 0)
                            {
                                worksheet.Cells["P" + Row].Value = 0;
                            }
                        }
                        shft++;
                        Row++;
                        Sno++;

                        worksheet.Row(Row).CustomHeight = false;

                        #endregion
                    }
                    UsedDateForExcel = UsedDateForExcel.AddDays(+1);
                }
            }

            //autofit
            //worksheet.Cells.AutoFitColumns(0);
            p.Save();
            //Downloding Excel
            string path1 = System.IO.Path.Combine(FileDir, "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx");
            System.IO.FileInfo file1 = new System.IO.FileInfo(path1);
            string Outgoingfile = "Utilization_Report" + frda.ToString("yyyy-MM-dd") + ".xlsx";
            if (file1.Exists)
            {
                //Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.AddHeader("Content-Disposition", "attachment; filename=" + Outgoingfile);
                //Response.AddHeader("Content-Length", file1.Length.ToString());
                //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                //Response.WriteFile(file1.FullName);
                //Response.Flush();
                //Response.Close();
            }
            //return View();
        }

        public async Task<int> GetOPIDleBreakDown(string CorrectedDate, int MachineID, string Colour)
        {
            DateTime currentdate = Convert.ToDateTime(CorrectedDate);
            string datetime = currentdate.ToString("yyyy-MM-dd");
            DataTable OP = new DataTable();
            int[] count = new int[4];
            using (MsqlConnection mc = new MsqlConnection())
            {
                //operating
                mc.open();
                String query1 = "SELECT count(ID) From [i_facility_tal].[dbo].tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(OP);
                mc.close();
            }
            if (OP.Rows.Count != 0)
            {
                count[0] = Convert.ToInt32(OP.Rows[0][0]);
            }
            return count[0];
        }

        public async Task<int> GetOPIDleBreakDownSetup(string CorrectedDate, int MachineID)
        {
            DateTime currentdate = Convert.ToDateTime(CorrectedDate);
            string datetime = currentdate.ToString("yyyy-MM-dd");
            string StartTime = null;
            string EndTime = null;

            var dayStartTiming = db.tbldaytimings.Where(m => m.IsDeleted == 0);
            foreach (var time in dayStartTiming)
            {
                StartTime = time.StartTime.ToString();
                EndTime = time.EndTime.ToString();
                break;
            }

            int[] count = new int[4];
            count[0] = 0;
            DataTable OP = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                //operating
                mc.open();
                String query1 = "SELECT StartDateTime,EndDateTime From [i_facility_tal].[dbo].tbllossofentry WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND MessageCodeID = 81";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(OP);
                mc.close();

                string StartTimedt = null;
                string EndTimedt = null;
                if (OP.Rows.Count != 0)
                {
                    for (int i = 0; i < OP.Rows.Count; i++)
                    {
                        DateTime stdt = DateTime.Now;
                        DateTime enddt = DateTime.Now;
                        if (string.IsNullOrEmpty(OP.Rows[i][0].ToString()) == false)
                        {
                            StartTimedt = OP.Rows[i][0].ToString();
                            stdt = Convert.ToDateTime(StartTimedt);
                            StartTimedt = stdt.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (string.IsNullOrEmpty(OP.Rows[i][1].ToString()) == false)
                        {
                            EndTimedt = OP.Rows[i][1].ToString();
                            enddt = Convert.ToDateTime(EndTimedt);
                            EndTimedt = enddt.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (StartTimedt != null && EndTimedt != null)
                        {
                            mc.open();
                            String query2 = "SELECT count(ID) From [i_facility_tal].[dbo].tbldailyprodstatus WHERE StartTime >='" + StartTimedt + "' AND EndTime <='" + EndTimedt + "' AND CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + "";
                            SqlDataAdapter da3 = new SqlDataAdapter(query2, mc.msqlConnection);
                            DataTable OP1 = new DataTable();
                            da3.Fill(OP1);
                            mc.close();
                            if (string.IsNullOrEmpty(OP1.Rows[0][0].ToString()) == false)
                                count[0] = count[0] + Convert.ToInt32(OP1.Rows[0][0]);
                        }
                    }
                }
            }
            return count[0];
        }

        public async Task<double> GetQuality(string CorrectedDate, int MachineID)
        {
            double[] count = new double[4];
            DataTable OP = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                //operating
                mc.open();
                String query1 = "SELECT Delivered_Qty,Rej_Qty From [i_facility_tal].[dbo].tblhmiscreen WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " ";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(OP);
                mc.close();
            }
            if (OP.Rows.Count != 0)
            {
                double Sumdelvrd = 0;
                double Sumrej = 0;
                for (int i = 0; i < OP.Rows.Count; i++)
                {
                    double delivdqty = 0;
                    if (string.IsNullOrEmpty(OP.Rows[i][0].ToString()) == false)
                        delivdqty = Convert.ToDouble(OP.Rows[i][0]);
                    double Rejqty = 0;
                    if (string.IsNullOrEmpty(OP.Rows[i][1].ToString()) == false)
                        Rejqty = Convert.ToDouble(OP.Rows[i][1]);
                    Sumdelvrd = Sumdelvrd + delivdqty;
                    Sumrej = Sumrej + Rejqty;
                }
                if (Sumdelvrd != 0)
                    count[0] = ((Sumdelvrd) / (Sumdelvrd + Sumrej));
                else
                    count[0] = ((Sumdelvrd - Sumrej) / 1);
            }
            return count[0];
        }

        //code to get the minor losses
        //take all rows in mode table and see if row's mode is idle,
        // then get difference between this row and next row's insertedOn time
        //if its <= 2min add it to minor loss......
        //public int GetMinorLoss(string CorrectedDate, int MachineID, string Colour)
        //{
        //    DateTime currentdate = Convert.ToDateTime(CorrectedDate);
        //    string datetime = currentdate.ToString("yyyy-MM-dd");

        //    //int[] count = new int[4];
        //    //MsqlConnection mc = new MsqlConnection();
        //    //mc.open();
        //    ////operating
        //    //mc.open();
        //    //var query1 = "SELECT * From tbldailyprodstatus WHERE CorrectedDate='" + CorrectedDate + "' AND MachineID=" + MachineID + " AND ColorCode='" + Colour + "'";
        //    int minorloss = 0;
        //    var modeData = db.tblmodes.Where(m => m.IsDeleted == 0 && m.MachineID == MachineID && m.CorrectedDate == CorrectedDate ).OrderBy(m => m.InsertedOn).ToList();
        //        for (int j = 0; j < modeData.Count(); j++)
        //        {
        //            if (j + 1 == modeData.Count()) //to insert last row of tblmode (b'cause of tblmodeData[j + 1] anamoly)
        //            {
        //               //have to takecare if shift ends with IDLE or just now got one yellow.
        //            }
        //            else
        //            {
        //                if (modeData[j].Mode == "IDLE")
        //                {
        //                    DateTime endTime = modeData[j+1].InsertedOn;
        //                    DateTime startTime = modeData[j].InsertedOn;
        //                    TimeSpan span = endTime.Subtract(startTime);
        //                    int Duration = Convert.ToInt32(span.Minutes);
        //                    if (Duration <= 2)
        //                    {
        //                        minorloss+= Duration;
        //                    }
        //                }
        //            }

        //        }
        //        //return minorloss;
        //}

        public async Task<int> GetMinorLoss(string CorrectedDate, int MachineID, string Colour)
        {
            DateTime currentdate = Convert.ToDateTime(CorrectedDate);
            string datetime = currentdate.ToString("yyyy-MM-dd");

            int minorloss = 0;
            using (i_facility_talEntities1 dbmode = new i_facility_talEntities1())
            {
                var modeData = dbmode.tblmodes.Where(m => m.IsDeleted == 0 && m.MachineID == MachineID && m.CorrectedDate == CorrectedDate).OrderBy(m => m.InsertedOn).ToList();
                for (int j = 0; j < modeData.Count(); j++)
                {
                    if (modeData[j].Mode == "IDLE")
                    {
                        DateTime endTime = modeData[j].InsertedOn;
                        DateTime startTime = modeData[j].InsertedOn;
                        TimeSpan span = endTime.Subtract(startTime);
                        int Duration = Convert.ToInt32(span.Minutes);
                        if (Duration <= 2)
                        {
                            minorloss += Duration;
                        }
                    }
                }
            }
            return minorloss;
        }

        //code to calculate scrap for the day.
        public async Task<double> GetScrap(string today, int machineid, double NOP)
        {
            double scrap = 0;
            using (i_facility_talEntities1 dbloss = new i_facility_talEntities1())
            {
                var scrapData = dbloss.tblhmiscreens.Where(m => m.CorrectedDate == today && m.MachineID == machineid && m.isWorkInProgress == 1);//isWorkInProgress: 1  work Completed
                double rejected = 0;
                double totalQuantity = 0;
                foreach (var row in scrapData)
                {
                    rejected += Convert.ToInt32(row.Rej_Qty);
                    totalQuantity += Convert.ToInt32(row.Target_Qty);
                }

                if (rejected != 0 && NOP != 0 && totalQuantity != 0)
                {
                    scrap = rejected * (NOP / totalQuantity);
                }
                else
                {
                    scrap = 0;
                }
            }
            return scrap;
        }

        public async Task<double> GetSettingTime(string UsedDateForExcel, int MachineID)
        {
            double settingTime = 0;
            int setupid = 0;
            string settingString = "Setup";
            using (i_facility_talEntities1 dbloss = new i_facility_talEntities1())
            {
                var setupiddata = dbloss.tbllossescodes.Where(m => m.IsDeleted == 0 && m.MessageType.Contains(settingString)).FirstOrDefault();
                if (setupiddata != null)
                {
                    setupid = setupiddata.LossCodeID;
                }
                else
                {
                    //Session["Error"] = "Unable to get Setup's ID";
                    //return -1;
                }
                //getting all setup's sublevels ids.
                var SettingIDs = dbloss.tbllossescodes.Where(m => m.IsDeleted == 0 && (m.LossCodesLevel1ID == setupid || m.LossCodesLevel2ID == setupid)).Select(m => m.LossCodeID).ToList();

                //settingTime = (from row in db.tbllossofentries
                //               where  row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );

                var SettingData = dbloss.tbllossofentries.Where(m => SettingIDs.Contains(m.MessageCodeID) && m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
                foreach (var row in SettingData)
                {
                    DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                    DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                    settingTime += endTime.Subtract(startTime).TotalMinutes;
                }
            }
            return settingTime;
        }

        public async Task<double> GetDownTimeLosses(string UsedDateForExcel, int MachineID, string contribute)
        {
            double LossTime = 0;
            //string contribute = "ROA";
            //getting all ROA sublevels ids. Only those of IDLE.
            using (i_facility_talEntities1 dbloss = new i_facility_talEntities1())
            {
                var SettingIDs = dbloss.tbllossescodes.Where(m => m.IsDeleted == 0 && m.ContributeTo == contribute && (m.MessageType != "PM" || m.MessageType != "BREAKDOWN")).Select(m => m.LossCodeID).ToList();

                //settingTime = (from row in db.tbllossofentries
                //               where  row.CorrectedDate == UsedDateForExcel && row.MachineID == MachineID );

                var SettingData = dbloss.tbllossofentries.Where(m => SettingIDs.Contains(m.MessageCodeID) && m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
                foreach (var row in SettingData)
                {
                    DateTime startTime = Convert.ToDateTime(row.StartDateTime);
                    DateTime endTime = Convert.ToDateTime(row.EndDateTime);
                    LossTime += endTime.Subtract(startTime).TotalMinutes;
                }
            }
            return LossTime;
        }

        public async Task<double> GetDownTimeBreakdown(string UsedDateForExcel, int MachineID)
        {
            double LossTime = 0;
            using (i_facility_talEntities1 dbbreak = new i_facility_talEntities1())
            {
                var BreakdownData = dbbreak.tblbreakdowns.Where(m => m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
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
            }
            return LossTime;
        }

        public async Task<double> GetSummationOfSCTvsPP(string UsedDateForExcel, int MachineID)
        {
            //if (MachineID == 14)
            //{
            //}
            double SummationofTime = 0;
            List<string> OccuredWOs = new List<string>();
            //To Extract Single WorkOrder Cutting Time
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                var PartsDataAll = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0 && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).OrderByDescending(m => m.HMIID).ToList();
                if (PartsDataAll.Count == 0)
                {
                    ////return SummationofTime;
                }
                foreach (var row in PartsDataAll)
                {
                    string partNo = row.PartNo;
                    string woNo = row.Work_Order_No;
                    string opNo = row.OperationNo;

                    string occuredwo = partNo + "," + woNo + "," + opNo;
                    if (!OccuredWOs.Contains(occuredwo))
                    {
                        OccuredWOs.Add(occuredwo);
                        var PartsData = dbhmi.tblhmiscreens.
                            Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0
                                && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)
                                && m.Work_Order_No == woNo && m.PartNo == partNo && m.OperationNo == opNo).
                                OrderByDescending(m => m.HMIID).ToList();

                        int totalpartproduced = 0;
                        int ProcessQty = 0, DeliveredQty = 0;
                        //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

                        #region new code

                        //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
                        int isHMIFirst = 2; //default NO History for that wo,pn,on

                        var mulitwoData = dbhmi.tbl_multiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWOID).Take(1).ToList();
                        //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.isWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

                        //Note: we are in this loop => hmiscreen table data is Available

                        if (mulitwoData.Count > 0)
                        {
                            isHMIFirst = 1;
                        }
                        else if (PartsData.Count > 0)
                        {
                            isHMIFirst = 0;
                        }
                        else if (PartsData.Count > 0 && mulitwoData.Count > 0) //for both Dates, now check for greatest amongst
                        {
                            int hmiIDFromMulitWO = row.HMIID;
                            DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
                                                                          where r.HMIID == hmiIDFromMulitWO
                                                                          select r.Time
                                                                          );
                            DateTime hmiDateTime = Convert.ToDateTime(row.Time);

                            if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
                            {
                                isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
                            }
                            else
                            {
                                isHMIFirst = 0;
                            }
                        }
                        if (isHMIFirst == 1)
                        {
                            string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
                            int.TryParse(delivString, out DeliveredQty);
                            string processString = Convert.ToString(mulitwoData[0].ProcessQty);
                            int.TryParse(processString, out ProcessQty);

                        }
                        else if (isHMIFirst == 0)//Take Data from HMI
                        {
                            string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
                            int.TryParse(delivString, out DeliveredQty);
                            string processString = Convert.ToString(PartsData[0].ProcessQty);
                            int.TryParse(processString, out ProcessQty);
                        }

                        #endregion

                        totalpartproduced = DeliveredQty + ProcessQty;

                        #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

                        int stdCuttingTime = 0;
                        var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                        if (stdcuttingTimeData != null)
                        {
                            int stdcuttingval = Convert.ToInt32(stdcuttingTimeData.StdCuttingTime);
                            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                            if (Unit == "Hrs")
                            {
                                stdCuttingTime = stdcuttingval * 60;
                            }
                            else //Unit is Minutes
                            {
                                stdCuttingTime = stdcuttingval;
                            }
                        }
                        #endregion

                        SummationofTime += stdCuttingTime * totalpartproduced;
                    }
                }
            }

            //To Extract Multi WorkOrder Cutting Time
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                var PartsDataAll = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 1 && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)).ToList();
                if (PartsDataAll.Count == 0)
                {
                    ////return SummationofTime;
                }
                foreach (var row in PartsDataAll)
                {
                    string partNo = row.PartNo;
                    string woNo = row.Work_Order_No;
                    string opNo = row.OperationNo;

                    string occuredwo = partNo + "," + woNo + "," + opNo;
                    if (!OccuredWOs.Contains(occuredwo))
                    {
                        OccuredWOs.Add(occuredwo);
                        var PartsData = dbhmi.tblhmiscreens.
                            Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && m.IsMultiWO == 0
                                && m.isWorkOrder == 0 && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0)
                                && m.Work_Order_No == woNo && m.PartNo == partNo && m.OperationNo == opNo).
                                OrderByDescending(m => m.HMIID).ToList();

                        int totalpartproduced = 0;
                        int ProcessQty = 0, DeliveredQty = 0;
                        //Decide to select deliveredQty & ProcessedQty lastest(from HMI or tblmultiWOselection)

                        #region new code

                        //here 1st get latest of delivered and processed among row in tblHMIScreen & tblmulitwoselection
                        int isHMIFirst = 2; //default NO History for that wo,pn,on

                        var mulitwoData = dbhmi.tbl_multiwoselection.Where(m => m.WorkOrder == woNo && m.PartNo == partNo && m.OperationNo == opNo).OrderByDescending(m => m.MultiWOID).Take(1).ToList();
                        //var hmiData = db.tblhmiscreens.Where(m => m.Work_Order_No == WONo && m.PartNo == Part && m.OperationNo == Operation && m.isWorkInProgress == 0).OrderByDescending(m => m.HMIID).Take(1).ToList();

                        //Note: we are in this loop => hmiscreen table data is Available

                        if (mulitwoData.Count > 0)
                        {
                            isHMIFirst = 1;
                        }
                        else if (PartsData.Count > 0)
                        {
                            isHMIFirst = 0;
                        }
                        else if (PartsData.Count > 0 && mulitwoData.Count > 0) //we have both Dates now check for greatest amongst
                        {
                            int hmiIDFromMulitWO = row.HMIID;
                            DateTime multiwoDateTime = Convert.ToDateTime(from r in db.tblhmiscreens
                                                                          where r.HMIID == hmiIDFromMulitWO
                                                                          select r.Time
                                                                          );
                            DateTime hmiDateTime = Convert.ToDateTime(row.Time);

                            if (Convert.ToInt32(multiwoDateTime.Subtract(hmiDateTime).TotalSeconds) > 0)
                            {
                                isHMIFirst = 1; // multiwoDateTime is greater than hmitable datetime
                            }
                            else
                            {
                                isHMIFirst = 0;
                            }
                        }

                        if (isHMIFirst == 1)
                        {
                            string delivString = Convert.ToString(mulitwoData[0].DeliveredQty);
                            int.TryParse(delivString, out DeliveredQty);
                            string processString = Convert.ToString(mulitwoData[0].ProcessQty);
                            int.TryParse(processString, out ProcessQty);
                        }
                        else if (isHMIFirst == 0) //Take Data from HMI
                        {
                            string delivString = Convert.ToString(PartsData[0].Delivered_Qty);
                            int.TryParse(delivString, out DeliveredQty);
                            string processString = Convert.ToString(PartsData[0].ProcessQty);
                            int.TryParse(processString, out ProcessQty);
                        }

                        #endregion

                        totalpartproduced = DeliveredQty + ProcessQty;

                        #region InnerLogic Common for both ways(HMI or tblmultiWOselection)

                        int stdCuttingTime = 0;
                        var stdcuttingTimeData = db.tblmasterparts_st_sw.Where(m => m.IsDeleted == 0 && m.OpNo == opNo && m.PartNo == partNo).FirstOrDefault();
                        if (stdcuttingTimeData != null)
                        {
                            int stdcuttingval = Convert.ToInt32(stdcuttingTimeData.StdCuttingTime);
                            string Unit = Convert.ToString(stdcuttingTimeData.StdCuttingTimeUnit);
                            if (Unit == "Hrs")
                            {
                                stdCuttingTime = stdcuttingval * 60;
                            }
                            else //Unit is Minutes
                            {
                                stdCuttingTime = stdcuttingval;
                            }
                        }
                        #endregion

                        SummationofTime += stdCuttingTime * totalpartproduced;
                    }
                }
            }
            return SummationofTime;
        }

        public async Task<int> GetMinorLosses(string CorrectedDate, int MachineID, string Colour)
        {
            DateTime currentdate = Convert.ToDateTime(CorrectedDate);
            string datetime = currentdate.ToString("yyyy-MM-dd");

            int minorloss = 0;
            int count = 0;
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                var Data = dbhmi.tbldailyprodstatus.Where(m => m.IsDeleted == 0 && m.MachineID == MachineID && m.CorrectedDate == CorrectedDate).OrderBy(m => m.StartTime).ToList();
                foreach (var row in Data)
                {
                    if (row.ColorCode == "yellow")
                    {
                        count++;
                    }
                    else
                    {
                        if (count > 0 && count < 2)
                        {
                            minorloss += count;
                            count = 0;

                        }
                        count = 0;
                    }
                }
            }
            return minorloss;
        }

        public async Task<double> GetScrapQtyTimeOfWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                var PartsData = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0) && m.isWorkOrder == 0).ToList();
                foreach (var row in PartsData)
                {
                    string partno = row.PartNo;
                    string operationno = row.OperationNo;
                    int scrapQty = Convert.ToInt32(row.Rej_Qty);
                    int DeliveredQty = Convert.ToInt32(row.Delivered_Qty);
                    DateTime startTime = Convert.ToDateTime(row.Date);
                    DateTime endTime = Convert.ToDateTime(row.Time);
                    //Double WODuration = endTimeTemp.Subtract(startTime).TotalMinutes;
                    Task<Double> WODuration = GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                    if ((scrapQty + DeliveredQty) == 0)
                    {
                        SQT += 0;
                    }
                    else
                    {
                        SQT += (Convert.ToDouble(WODuration) / (scrapQty + DeliveredQty)) * scrapQty;
                    }
                }
            }
            return SQT;
        }

        //GOD
        public async Task<double> GetScrapQtyTimeOfRWO(string UsedDateForExcel, int MachineID)
        {
            double SQT = 0;
            using (i_facility_talEntities1 dbhmi = new i_facility_talEntities1())
            {
                var PartsData = dbhmi.tblhmiscreens.Where(m => m.CorrectedDate == UsedDateForExcel && m.MachineID == MachineID && (m.isWorkInProgress == 1 || m.isWorkInProgress == 0) && m.isWorkOrder == 1).ToList();
                foreach (var row in PartsData)
                {
                    string partno = row.PartNo;
                    string operationno = row.OperationNo;
                    int scrapQty = Convert.ToInt32(row.Rej_Qty);
                    int DeliveredQty = Convert.ToInt32(row.Delivered_Qty);
                    DateTime startTime = Convert.ToDateTime(row.Date);
                    DateTime endTime = Convert.ToDateTime(row.Time);
                    Task<Double> WODuration = GetGreen(UsedDateForExcel, startTime, endTime, MachineID);

                    //Double WODuration = endTime.Subtract(startTime).TotalMinutes;
                    //For Availability Loss
                    //double Settingtime = GetSetupForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                    //double green = GetOT(UsedDateForExcel, startTime, endTime, MachineID);
                    //double DownTime = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROA");
                    //double BreakdownTime = GetBreakDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID);
                    //double AL = DownTime + BreakdownTime + Settingtime;

                    //For Performance Loss
                    //double downtimeROP = GetDownTimeForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "ROP");
                    //double minorlossWO = GetMinorLossForReworkLoss(UsedDateForExcel, startTime, endTime, MachineID, "yellow");
                    //double PL = downtimeROP + minorlossWO;

                    SQT += Convert.ToDouble(WODuration);
                }
            }
            return SQT;
        }

        #region
        public async Task<double> GetSetupForReworkLoss(string UsedDateForExcel, DateTime TSstartTime, DateTime TSendTime, int MachineID)
        {
            string setup = "Setup";
            double settingTime = 0;
            int setupid = 0;
            string settingString = "Setup";

            DateTime WOstarttimeDate = Convert.ToDateTime(TSstartTime);
            DateTime WOendtimeDate = Convert.ToDateTime(TSendTime);
            var setupiddata = db.tbllossescodes.Where(m => m.IsDeleted == 0 && m.MessageType.Contains(settingString)).FirstOrDefault();
            if (setupiddata != null)
            {
                setupid = setupiddata.LossCodeID;
            }
            else
            {
                //Session["Error"] = "Unable to get Setup's ID";
                //return -1;
            }
            //getting all setup's sublevels ids.
            //var SettingIDs = db.tbllossescodes.Where(m => m.IsDeleted == 0 && (m.LossCodesLevel1ID == setupid || m.LossCodesLevel2ID == setupid)).Select(m=>m.LossCodeID).ToList();

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT * From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and DoneWithRow = 1"
                    + " and (( StartDateTime <  '" + WOendtimeDate + " ' and  StartDateTime > '" + WOstarttimeDate + "') or "
                    + " ( EndDateTime >  '" + WOstarttimeDate + " ' and  EndDateTime < '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime <  '" + WOstarttimeDate + " ' and  EndDateTime < '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime >  '" + WOstarttimeDate + " ' and  StartDateTime < '" + WOendtimeDate + "'  and  EndDateTime > '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime <  '" + WOstarttimeDate + " ' and  EndDateTime > '" + WOendtimeDate + "' ))"
                    + " and MessageCodeID IN ( Select LossCodeID from [i_facility_tal].[dbo].tbllossescodes where MessageType = '" + setup + "')";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                //2 & 3
                DateTime Istarttime = Convert.ToDateTime(lossesData.Rows[0][2]);
                DateTime Iendtime = Convert.ToDateTime(lossesData.Rows[0][3]);

                double diffInMinutes = 0;
                if (Istarttime < WOstarttimeDate)
                {
                    //this Implies : Consider from wostarttimedate
                    if (Iendtime < WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                }
                else
                {
                    //this Implies : Consider from Istarttime
                    if (Iendtime <= WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(Istarttime).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(Istarttime).TotalMinutes;
                    }
                }
                settingTime += diffInMinutes;
            }
            return settingTime;
        }

        public async Task<double> GetDownTimeForReworkLoss(string UsedDateForExcel, DateTime TSstartTime, DateTime TSendTime, int MachineID, string roa)
        {

            double settingTime = 0;

            DateTime WOstarttimeDate = Convert.ToDateTime(TSstartTime);
            DateTime WOendtimeDate = Convert.ToDateTime(TSendTime);

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT * From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and DoneWithRow = 1"
                    + " and (( StartDateTime <  '" + WOendtimeDate + " ' and  StartDateTime > '" + WOstarttimeDate + "') or "
                    + " ( EndDateTime >  '" + WOstarttimeDate + " ' and  EndDateTime < '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime <  '" + WOstarttimeDate + " ' and  EndDateTime < '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime >  '" + WOstarttimeDate + " ' and  StartDateTime < '" + WOendtimeDate + "'  and  EndDateTime > '" + WOendtimeDate + "') or  "
                    + " ( StartDateTime <  '" + WOstarttimeDate + " ' and  EndDateTime > '" + WOendtimeDate + "' ))"
                    + " and MessageCodeID IN ( Select LossCodeID from [i_facility_tal].[dbo].tbllossescodes where ContributeTo = '" + roa + "')";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                //2 & 3
                DateTime Istarttime = Convert.ToDateTime(lossesData.Rows[0][2]);
                DateTime Iendtime = Convert.ToDateTime(lossesData.Rows[0][3]);

                double diffInMinutes = 0;
                if (Istarttime < WOstarttimeDate)
                {
                    //this Implies : Consider from wostarttimedate
                    if (Iendtime < WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                }
                else
                {
                    //this Implies : Consider from Istarttime
                    if (Iendtime <= WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(Istarttime).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(Istarttime).TotalMinutes;
                    }
                }
                settingTime += diffInMinutes;
            }
            return settingTime;
        }

        public async Task<double> GetBreakDownTimeForReworkLoss(string UsedDateForExcel, DateTime TSstartTime, DateTime TSendTime, int MachineID)
        {
            double BreakdownTime = 0;

            DateTime WOstarttimeDate = Convert.ToDateTime(TSstartTime);
            DateTime WOendtimeDate = Convert.ToDateTime(TSendTime);

            DataTable BreaddownData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT * From [i_facility_tal].[dbo].tblbreakdown WHERE MachineID = '" + MachineID + "' and DoneWithRow = 1 and CorrectedDate = '" + UsedDateForExcel + "'"
                    + " and (( StartTime <  '" + WOendtimeDate + " ' and  StartTime > '" + WOstarttimeDate + "') or "
                    + " ( EndTime >  '" + WOstarttimeDate + " ' and  EndTime < '" + WOendtimeDate + "') or  "
                    + " ( StartTime <  '" + WOstarttimeDate + " ' and  EndTime < '" + WOendtimeDate + "') or  "
                    + " ( StartTime >  '" + WOstarttimeDate + " ' and  StartTime < '" + WOendtimeDate + "'  and  EndTime > '" + WOendtimeDate + "') or  "
                    + " ( StartTime <  '" + WOstarttimeDate + " ' and  EndTime > '" + WOendtimeDate + "' ))"
                    + " and BreakDownCode IN ( Select BreakDownCodeID from [i_facility_tal].[dbo].tblbreakdowncodes where  EndTime <> '' )";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(BreaddownData);
                mc.close();
            }

            for (int i = 0; i < BreaddownData.Rows.Count; i++)
            {
                //2 & 3
                DateTime Istarttime = Convert.ToDateTime(BreaddownData.Rows[0][1]);
                DateTime Iendtime = Convert.ToDateTime(BreaddownData.Rows[0][2]);

                double diffInMinutes = 0;
                if (Istarttime < WOstarttimeDate)
                {
                    //this Implies : Consider from wostarttimedate
                    if (Iendtime < WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(WOstarttimeDate).TotalMinutes;
                    }
                }
                else
                {
                    //this Implies : Consider from Istarttime
                    if (Iendtime <= WOendtimeDate)
                    {
                        //this Implies : Consider till(including) Iendtime
                        diffInMinutes = Iendtime.Subtract(Istarttime).TotalMinutes;
                    }
                    else
                    {
                        //this Implies : Consider till(including ) woendtimedate
                        diffInMinutes = WOendtimeDate.Subtract(Istarttime).TotalMinutes;
                    }
                }
                BreakdownTime += diffInMinutes;
            }
            return BreakdownTime;
        }

        public async Task<int> GetMinorLossForReworkLoss(string CorrectedDate, DateTime TSstartTime, DateTime TSendTime, int MachineID, string Colour)
        {

            DateTime WOstarttimeDate = Convert.ToDateTime(TSstartTime);
            DateTime WOendtimeDate = Convert.ToDateTime(TSendTime);

            int minorloss = 0;
            int count = 0;

            DataTable yellowData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT * FROM [i_facility_tal].[dbo].tbldailyprodstatus where IsDeleted = 0 and MachineID = '" + MachineID + "' and CorrectedDate = '" + CorrectedDate + "' order by InsertedOn";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(yellowData);
                mc.close();
            }

            for (int i = 0; i < yellowData.Rows.Count; i++)
            {
                if (Convert.ToString(yellowData.Rows[0][1]) == "yellow")
                {
                    count++;
                }
                else
                {
                    if (count > 0 && count <= 2)
                    {
                        minorloss += count;
                        count = 0;

                    }
                    count = 0;
                }
            }
            return minorloss;
        }
        #endregion

        //Output in Minutes
        public async Task<double> GetGreen(string UsedDateForExcel, DateTime TSstartTime, DateTime TSendTime, int MachineID)
        {
            double OperatingTime = 0;
            double FirstRowDur = 0;
            double LastRowDur = 0;
            try
            {
                DataTable GreenData = new DataTable();
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query = "SELECT * From [i_facility_tal].[dbo].tblmode WHERE MachineID = '" + MachineID + "' and ColorCode = 'green' and IsCompleted = 1  and "
                   + "( StartTime <= '" + TSstartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + TSstartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + TSendTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + TSendTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                   + " ( StartTime > '" + TSstartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + TSendTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";

                    SqlDataAdapter da1 = new SqlDataAdapter(query, mc.msqlConnection);
                    da1.Fill(GreenData);
                    mc.close();
                }

                for (int i = 0; i < GreenData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(GreenData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(GreenData.Rows[i][1])))
                    {
                        DateTime LStartDate = Convert.ToDateTime(GreenData.Rows[i][9]);
                        DateTime LEndDate = Convert.ToDateTime(GreenData.Rows[i][10]);
                        double IndividualDur = Convert.ToDouble(GreenData.Rows[i][13]);

                        //Get Duration Based on start & end Time.

                        if (LStartDate < TSstartTime)
                        {
                            double StartDurationExtra = TSstartTime.Subtract(LStartDate).TotalSeconds;
                            IndividualDur -= StartDurationExtra;
                        }
                        if (LEndDate > TSendTime)
                        {
                            double EndDurationExtra = LEndDate.Subtract(TSendTime).TotalSeconds;
                            IndividualDur -= EndDurationExtra;
                        }
                        OperatingTime += IndividualDur;
                    }
                }
            }
            catch (Exception e)
            { }
            return Math.Round(OperatingTime / 60, 2);

        }

        public string TimeFromSeconds(double seconds)
        {
            string FinalTime = null;
            //int Days = Convert.ToInt32((seconds / (24 * 60 * 60)));
            //int Hours = Convert.ToInt32((seconds / 3600) % 24);// hours to sec 60 * 60 = 3600
            //int Minutes = Convert.ToInt32((seconds / 60) % 60);
            //int Seconds = Convert.ToInt32(seconds % 60);


            //string DaysString = null, HoursString = null, MinutesString = null, SecondsString = null;

            //DaysString = Days < 10 ? Convert.ToString("0" + Days) : Convert.ToString(Days);
            //HoursString = Hours < 10 ? Convert.ToString("0" + Hours) : Convert.ToString(Hours);
            //MinutesString = Minutes < 10 ? Convert.ToString("0" + Minutes) : Convert.ToString(Minutes);
            //SecondsString = Seconds < 10 ? Convert.ToString("0" + Seconds) : Convert.ToString(Seconds);

            //FinalTime = DaysString + ":" + HoursString + ":" + MinutesString + ":" + SecondsString;


            //here backslash is must to tell that colon is
            //not the part of format, it just a character that we want in output
            TimeSpan t = TimeSpan.FromSeconds(seconds);

            if (t.Days > 0)
            {
                FinalTime = t.ToString(@"d\D\,\ hh\:mm\:ss");
            }
            else
            {
                FinalTime = t.ToString(@"\0\D\,\ hh\:mm\:ss");
            }
            return FinalTime;
        }

        public async Task<string> LossHierarchy(int LossCodeID)
        {
            //string losshierarchy = LossCode + " : ";
            string losshierarchy = null;
            var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossCodeID).FirstOrDefault();

            int level = lossdata.LossCodesLevel;
            if (level == 1)
            {
                if (LossCodeID == 999)
                {
                    losshierarchy += "NoCode Entered";
                }
                else
                {
                    losshierarchy += lossdata.LossCode;
                }
            }
            else if (level == 2)
            {
                int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

                losshierarchy += lossdata1.LossCode + "->" + lossdata.LossCode;
            }
            else if (level == 3)
            {
                int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                losshierarchy += lossdata1.LossCode + "->" + lossdata2.LossCode + "->" + lossdata.LossCode;
            }
            return losshierarchy;

        }

        public async Task<double> GetAnyLossDurationInSec(int machineID, int lossId, string CorrectedDate)
        {
            double duration = 0;
            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT StartDateTime,EndDateTime From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + machineID + "' and CorrectedDate = '" + CorrectedDate + "' and MessageCodeID = '" + lossId + "' and DoneWithRow = 1";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }
            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime StartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime EndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
                    duration += EndDate.Subtract(StartDate).TotalSeconds;
                }
            }
            return duration;
        }

        public List<KeyValuePair<int, double>> GetAllLossesDurationSeconds(int machineID, string CorrectedDate)
        {
            List<KeyValuePair<int, double>> durationList = new List<KeyValuePair<int, double>>();
            DataTable lossesData = new DataTable();
            var LossesIDs = db.tbllossofentries.Where(m => m.MachineID == machineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeID).Distinct().ToList();
            foreach (var loss in LossesIDs)
            {
                lossesData.Clear();
                double duration = 0;
                int lossID = loss;
                //if (lossID == 2)
                //{
                //}
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + machineID + "' and CorrectedDate = '" + CorrectedDate + "' and MessageCodeID = '" + lossID + "' and DoneWithRow = 1 ";
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(lossesData);
                    mc.close();
                }

                for (int i = 0; i < lossesData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                    {
                        DateTime StartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                        DateTime EndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
                        duration += EndDate.Subtract(StartDate).TotalSeconds;
                    }
                }
                durationList.Add(new KeyValuePair<int, double>(lossID, duration));
            }
            return durationList;
        }

        public List<KeyValuePair<int, double>> GetAllBreakdownDurationSeconds(int machineID, string CorrectedDate)
        {
            List<KeyValuePair<int, double>> durationList = new List<KeyValuePair<int, double>>();
            var LossesIDs = db.tblbreakdowns.Where(m => m.MachineID == machineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.BreakDownCode).Distinct().ToList();
            foreach (var loss in LossesIDs)
            {
                DataTable lossesData = new DataTable();
                double duration = 0;
                int lossID = Convert.ToInt32(loss);
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query1 = "SELECT StartTime,EndTime From [i_facility_tal].[dbo].tblbreakdown WHERE MachineID = '" + machineID + "' and CorrectedDate = '" + CorrectedDate + "' and BreakDownCode = '" + lossID + "' and DoneWithRow = 1 ";
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(lossesData);
                    mc.close();
                }

                for (int i = 0; i < lossesData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                    {
                        DateTime StartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                        DateTime EndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
                        duration += EndDate.Subtract(StartDate).TotalSeconds;
                    }
                }
                duration = duration;
                durationList.Add(new KeyValuePair<int, double>(lossID, duration));
            }
            return durationList;
        }

        public List<KeyValuePair<int, double>> GetAllLossesDurationSecondsForWO(int machineID, string CorrectedDate, DateTime StartTime, DateTime EndTime)
        {
            List<KeyValuePair<int, double>> durationList = new List<KeyValuePair<int, double>>();
            var LossesIDs = db.tbllossofentries.Where(m => m.MachineID == machineID && m.CorrectedDate == CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeID).Distinct().ToList();
            foreach (var loss in LossesIDs)
            {
                DataTable lossesData = new DataTable();
                double duration = 0;
                int lossID = loss;
                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + machineID + "' and CorrectedDate = '" + CorrectedDate + "' and MessageCodeID = '" + lossID + "' and DoneWithRow = 1  and "
                        + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                        + " (  StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(lossesData);
                    mc.close();
                }

                for (int i = 0; i < lossesData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                    {
                        DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                        DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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
                durationList.Add(new KeyValuePair<int, double>(lossID, duration));
            }
            return durationList;
        }

        //Output: In Seconds
        public async Task<double> GetSettingTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double settingTime = 0;
            int setupid = 0;
            string settingString = "Setup";
            var setupiddata = db.tbllossescodes.Where(m => m.IsDeleted == 0 && m.MessageType.Equals(settingString, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (setupiddata != null)
            {
                setupid = setupiddata.LossCodeID;
            }
            else
            {
                //Session["Error"] = "Unable to get Setup's ID";
                //return -1;
            }

            // var s = string.Join(",", products.Where(p => p.ProductType == someType).Select(p => p.ProductId.ToString()));
            //getting all setup's sublevels ids.
            var SettingIDs = db.tbllossescodes
                                .Where(m => m.LossCodesLevel1ID == setupid)
                                .Select(m => m.LossCodeID).ToList()
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
            //var SettingData = db.tbllossofentries.Where(m => SettingIDs.Contains(m.MessageCodeID) && m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
            //foreach (var row in SettingData)
            //{
            //    DateTime startTime = Convert.ToDateTime(row.StartDateTime);
            //    DateTime endTime = Convert.ToDateTime(row.EndDateTime);
            //    settingTime += endTime.Subtract(startTime).TotalMinutes;
            //}

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and MessageCodeID IN ( " + SettingIDsString + " ) and DoneWithRow = 1  and "
                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                    + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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

            return settingTime;
        }
        //Output: In Seconds
        public async Task<double> GetChangeoverTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double ChangeOverTime = 0;
            int ChangeOverid = 0;
            string ChangeOverString = "Changeover";
            //var ChangeOverdata = db.tbllossescodes.Where(m => m.MessageType.Equals(ChangeOverString, StringComparison.OrdinalIgnoreCase) && m.IsDeleted == 0  && m.LossCodesLevel1ID == 1).FirstOrDefault();
            var ChangeOverdata = db.tbllossescodes.Where(m => m.MessageType.Equals(ChangeOverString, StringComparison.OrdinalIgnoreCase) && m.IsDeleted == 0 && m.LossCodesLevel1ID == 1).FirstOrDefault();
            if (ChangeOverdata != null)
            {
                ChangeOverid = ChangeOverdata.LossCodeID;
            }
            else
            {
                //Session["Error"] = "Unable to get ChangeOver ID";
                //return -1;
            }
            var ChangeOverIDs = db.tbllossescodes
                                .Where(m => m.LossCodesLevel1ID == ChangeOverid)
                                .Select(m => m.LossCodeID).ToList()
                                .Distinct();
            string ChangeOverIDsString = null;
            int j = 0;
            foreach (var row in ChangeOverIDs)
            {
                if (j != 0)
                {
                    ChangeOverIDsString += "," + Convert.ToInt32(row);
                }
                else
                {
                    ChangeOverIDsString = Convert.ToInt32(row).ToString();
                }
                j++;
            }

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and MessageCodeID IN ( " + ChangeOverIDsString + " ) and DoneWithRow = 1  and "
                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                    + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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
                    ChangeOverTime += IndividualDur;
                }
            }

            return ChangeOverTime;
        }
        //Output: In Seconds
        public async Task<double> GetAllLossesTimeForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double AllLossesTime = 0;
            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                // String query1 = "SELECT StartDateTime,EndDateTime,LossID From tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and DoneWithRow = 1  and "
                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and DoneWithRow = 1  and "
                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                    + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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
            }

            return AllLossesTime;
        }
        //Output: In Seconds
        public async Task<double> GetDownTimeBreakdownForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double BreakdownTime = 0;

            //var BreakdownData = db.tblbreakdowns.Where(m => m.MachineID == MachineID && m.CorrectedDate == UsedDateForExcel && m.DoneWithRow == 1).ToList();
            //foreach (var row in BreakdownData)
            //{
            //    if ((Convert.ToString(row.EndTime) == null) || row.EndTime == null)
            //    {
            //        //do nothing
            //    }
            //    else
            //    {
            //        DateTime startTime = Convert.ToDateTime(row.StartTime);
            //        DateTime endTime = Convert.ToDateTime(row.EndTime);
            //        BreakdownTime += endTime.Subtract(startTime).TotalMinutes;
            //    }
            //}

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT StartTime,EndTime From [i_facility_tal].[dbo].tblbreakdown WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and DoneWithRow = 1  and "
                    + "( StartTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                    + " ( StartTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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
                    BreakdownTime += IndividualDur;
                }
            }

            return BreakdownTime;
        }

        // Output: In Seconds
        public async Task<double> GetSelfInsepectionForWO(string UsedDateForExcel, int MachineID, DateTime StartTime, DateTime EndTime)
        {
            double SelfInspectionTime = 0;
            int SelfInspectionid = 112;

            DataTable lossesData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + MachineID + "' and CorrectedDate = '" + UsedDateForExcel + "' and MessageCodeID IN ( " + SelfInspectionid + " ) and DoneWithRow = 1  and "
                    + "( StartDateTime <= '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( ( EndDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( EndDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' or   EndDateTime > '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' )  ) ) or "
                    + " ( StartDateTime > '" + StartTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and ( StartDateTime < '" + EndTime.ToString("yyyy-MM-dd HH:mm:ss") + "' ) ))";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesData);
                mc.close();
            }

            for (int i = 0; i < lossesData.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                {
                    DateTime LStartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                    DateTime LEndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
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
            }

            return SelfInspectionTime;
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

        public async Task<string> LossHierarchy3rdLevel(int LossCode)
        {
            string losshierarchy = LossCode + " : ";
            var lossdata = db.tbllossescodes.Where(m => m.LossCodeID == LossCode).FirstOrDefault();

            int level = lossdata.LossCodesLevel;
            //if (level == 1)
            //{
            //    if (LossCode == 999)
            //    {
            //        losshierarchy += "NoCode Entered";
            //    }
            //    else
            //    {
            //        losshierarchy += lossdata.LossCode;
            //    }
            //}
            //else if (level == 2)
            //{
            //    int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
            //    var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();

            //    losshierarchy += lossdata1.LossCode + ":" + lossdata.LossCode;
            //}
            //else 
            if (level == 3)
            {
                int lossLevel1ID = Convert.ToInt32(lossdata.LossCodesLevel1ID);
                int lossLevel2ID = Convert.ToInt32(lossdata.LossCodesLevel2ID);
                var lossdata1 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel1ID).FirstOrDefault();
                var lossdata2 = db.tbllossescodes.Where(m => m.LossCodeID == lossLevel2ID).FirstOrDefault();
                losshierarchy += lossdata1.LossCode + " :: " + lossdata2.LossCode + " : " + lossdata.LossCode;
            }
            return losshierarchy;
        }

        //HoldHierarchy 2017-03-23
        string GetHoldHierarchy(int holdId)
        {
            string HoldHierarchy = null;
            var holdData = db.tblholdcodes.Where(m => m.HoldCodeID == holdId).FirstOrDefault();

            int level = holdData.HoldCodesLevel;
            if (level == 1)
            {

                HoldHierarchy += holdData.HoldCode;
            }
            else if (level == 2)
            {
                int lossLevel1ID = Convert.ToInt32(holdData.HoldCodesLevel1ID);
                var lossdata1 = db.tblholdcodes.Where(m => m.HoldCodeID == lossLevel1ID).FirstOrDefault();

                HoldHierarchy += lossdata1.HoldCode + "->" + holdData.HoldCode;
            }
            else if (level == 3)
            {
                int lossLevel1ID = Convert.ToInt32(holdData.HoldCodesLevel1ID);
                int lossLevel2ID = Convert.ToInt32(holdData.HoldCodesLevel2ID);
                var lossdata1 = db.tblholdcodes.Where(m => m.HoldCodeID == lossLevel1ID).FirstOrDefault();
                var lossdata2 = db.tblholdcodes.Where(m => m.HoldCodeID == lossLevel2ID).FirstOrDefault();
                HoldHierarchy += lossdata1.HoldCode + "->" + lossdata2.HoldCode + "->" + holdData.HoldCode;
            }
            return HoldHierarchy;
        }

        public void IntoFile(string Msg)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Application.StartupPath + @"\AutoReportMailLogFile.txt";
            using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
            {
                writer.WriteLine(System.DateTime.Now + ":  " + Msg);
            }

        }

        public async Task<bool> sendMail(int AutoReportID, DateTime StartTimeFromEscLogTable, string PathWithFileNameAndExtension)
        {
            bool Status = false;
            try
            {
                DataTable dtEscData = new DataTable();
                MsqlConnection mcEscData = new MsqlConnection();
                mcEscData.open();
                String queryEscData = "SELECT ToMailList,CCMailList,PlantID,ShopID,CellID,MachineID,ReportID From [i_facility_tal].[dbo].tbl_autoreportsetting WHERE AutoReportID = " + AutoReportID;
                SqlDataAdapter daEscData = new SqlDataAdapter(queryEscData, mcEscData.msqlConnection);
                daEscData.Fill(dtEscData);
                mcEscData.close();

                string toMailIDs = dtEscData.Rows[0][0].ToString();
                string ccMailIDs = dtEscData.Rows[0][1].ToString();
                string PlantId = Convert.ToString(dtEscData.Rows[0][2]);
                string ShopId = Convert.ToString(dtEscData.Rows[0][3]);
                string CellId = Convert.ToString(dtEscData.Rows[0][4]);
                string WCId = Convert.ToString(dtEscData.Rows[0][5]);
                int reportID = Convert.ToInt32(dtEscData.Rows[0][6]);

                Task<string> Hierarchy = HierarchyPSCM(PlantId, ShopId, CellId, WCId);
                string ReportName = Convert.ToString(db.tbl_reportmaster.Where(m => m.ReportID == reportID).Select(m => m.ReportDispName).FirstOrDefault());

                MailMessage mail = new MailMessage();

                string[] Tomails = toMailIDs.Split(',');
                foreach (var mailid in Tomails)
                {
                    string mailID = Convert.ToString(mailid).Trim();
                    if (!string.IsNullOrEmpty(mailID))
                    {
                        mail.To.Add(new MailAddress(mailid));
                    }
                }

                string[] Ccmails = ccMailIDs.Split(',');
                foreach (var mailid in Ccmails)
                {
                    string mailID = Convert.ToString(mailid).Trim();
                    if (!string.IsNullOrEmpty(mailID))
                    {
                        mail.CC.Add(new MailAddress(mailid));
                    }
                }
                //mail.From = new MailAddress("unitworks@tasl.aero");
                mail.From = new MailAddress("barcodesupport@tal.co.in");
                mail.Subject = "Auto Report for " + ReportName;
                mail.Attachments.Add(new Attachment(PathWithFileNameAndExtension));

                mail.Body = "<p><b>Dear Concerned,</b></p>" +
                            "<p><b></b></p>" +
                            "<p><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; This is Automatic Report for  </b></p>" +
                            "<b> " + Hierarchy.Result + "</b>" +
                    // "<table><tr> <td> " + ReasonPath + "</td> <td> From " + StartTimeFromEscLogTable + " .</td>  <td> Escalated after " + Duration + " Minutes</td></tr></table>" +
                             "<p><b></b></p>" +
                             "<p><b></b></p>" +
                             "<p><b> Regards,</b></p>" +
                            "<p><b> TAL - DAS team</b></p>" +

                             "<p><b></b></p>" +
                             "<p><b></b></p>" +
                            "<p><b>Note: This is an autogenerated E-Mail for the Automatic Reports. Do Not Reply back on this Mail.</b></p>" +
                            "<p><b></b></p>";

                mail.IsBodyHtml = true;
                SmtpClient smtp = new SmtpClient();
                //smtp.Host = "10.30.10.57";
                smtp.Host = MsqlConnection.Host;  //server
                smtp.Port = MsqlConnection.Portmail;
                //smtp.Port = 587;
                smtp.UseDefaultCredentials = false;

                //smtp.Credentials = CredentialCache.DefaultNetworkCredentials;
                //smtp.Credentials = new System.Net.NetworkCredential("unitworks@tasl.aero", "abcd@1234", "tasl.aero");  //server
                smtp.Credentials = new System.Net.NetworkCredential(MsqlConnection.UserNamemail, MsqlConnection.Passwordmail, MsqlConnection.Domain);   //server
                //smtp.Credentials = new System.Net.NetworkCredential("unitworks@tasl.aero", "527293");
                //smtp.EnableSsl = false;
                //smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.EnableSsl = true;
                smtp.Send(mail);
                Status = true;
                // MessageBox.Show("Status : success ");
            }
            catch (Exception e)
            {
                // MessageBox.Show("Error :: sendMail : " + e);
                Status = false;
                IntoFile("Error While Sending Mail " + e);
            }
            return await Task.FromResult<bool> (Status);
        }

        public async Task<string> HierarchyPSCM(string plant, string shop = null, string cell = null, string mac = null)
        {
            string Hierarchy = null;
            int PlantID = 0, ShopID = 0, CellID = 0, WCID = 0;
            int.TryParse(plant, out PlantID);
            int.TryParse(shop, out ShopID);
            int.TryParse(cell, out CellID);
            int.TryParse(mac, out WCID);
            string PlantName = null, ShopName = null, CellName = null, WCName = null;

            if (PlantID != 0)
            {
                PlantName = Convert.ToString(db.tblplants.Where(m => m.PlantID == PlantID).Select(m => m.PlantName).SingleOrDefault());
                if (PlantName != null)
                {
                    Hierarchy = PlantName;
                }
            }
            if (ShopID != 0)
            {
                ShopName = Convert.ToString(db.tblshops.Where(m => m.ShopID == ShopID).Select(m => m.ShopName).SingleOrDefault());
                if (ShopName != null)
                {
                    Hierarchy += "-->" + ShopName;
                }
            }
            if (CellID != 0)
            {
                CellName = Convert.ToString(db.tblcells.Where(m => m.CellID == CellID).Select(m => m.CellName).SingleOrDefault());
                if (CellName != null)
                {
                    Hierarchy += "-->" + CellName;
                }
            }
            if (WCID != 0)
            {
                WCName = Convert.ToString(db.tblmachinedetails.Where(m => m.MachineID == WCID).Select(m => m.MachineDispName).SingleOrDefault());
                if (WCName != null)
                {
                    Hierarchy += "-->" + WCName;
                }
            }

            return Hierarchy;
        }

        private static void GetWeek(DateTime now, CultureInfo cultureInfo, out DateTime begining, out DateTime end)
        {
            if (now == null)
                throw new ArgumentNullException("now");
            if (cultureInfo == null)
                throw new ArgumentNullException("cultureInfo");

            var firstDayOfWeek = cultureInfo.DateTimeFormat.FirstDayOfWeek;
            int offset = firstDayOfWeek - now.DayOfWeek;
            if (offset != 1)
            {
                DateTime weekStart = now.AddDays(offset);
                DateTime endOfWeek = weekStart.AddDays(6);
                begining = weekStart;
                end = endOfWeek;
            }
            else
            {
                begining = now.AddDays(-6);
                end = now;
            }
        }

        public List<KeyValuePair<int, double>> GetAllLossesDurationSecondsForDateRange(int machineID, string CorrectedDateS, string CorrectedDateE)
        {
            List<KeyValuePair<int, double>> durationList = new List<KeyValuePair<int, double>>();
            DataTable lossesData = new DataTable();
            DataTable lossesIDData = new DataTable();
            using (MsqlConnection mc = new MsqlConnection())
            {
                mc.open();
                String query1 = "SELECT distinct( MessageCodeID ) From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + machineID + "' and CorrectedDate >= '" + CorrectedDateS + "' and CorrectedDate <= '" + CorrectedDateE + "' and DoneWithRow = 1 ";
                SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                da1.Fill(lossesIDData);
                mc.close();
            }
            //var LossesIDs = db.tbllossofentries.Where(m => m.MachineID == machineID && m.CorrectedDate >= CorrectedDate && m.DoneWithRow == 1).Select(m => m.MessageCodeID).Distinct().ToList();
            for (int j = 0; j < lossesIDData.Rows.Count; j++)
            {
                lossesData.Clear();
                double duration = 0;
                int lossID = Convert.ToInt32(lossesIDData.Rows[j][0]);

                using (MsqlConnection mc = new MsqlConnection())
                {
                    mc.open();
                    String query1 = "SELECT StartDateTime,EndDateTime,LossID From [i_facility_tal].[dbo].tbllossofentry WHERE MachineID = '" + machineID + "' and CorrectedDate >= '" + CorrectedDateS + "' and CorrectedDate <= '" + CorrectedDateE + "' and MessageCodeID = '" + lossID + "' and DoneWithRow = 1 ";
                    SqlDataAdapter da1 = new SqlDataAdapter(query1, mc.msqlConnection);
                    da1.Fill(lossesData);
                    mc.close();
                }
                for (int i = 0; i < lossesData.Rows.Count; i++)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][0])) && !string.IsNullOrEmpty(Convert.ToString(lossesData.Rows[i][1])))
                    {
                        DateTime StartDate = Convert.ToDateTime(lossesData.Rows[i][0]);
                        DateTime EndDate = Convert.ToDateTime(lossesData.Rows[i][1]);
                        duration += EndDate.Subtract(StartDate).TotalSeconds;
                    }
                }
                durationList.Add(new KeyValuePair<int, double>(lossID, duration));
            }
            return durationList;
        }

    }
}
