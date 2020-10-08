using i_facility_IdleHandlerWithOptimization.TataModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace i_facility_IdleHandlerWithOptimization
{
    public partial class Form1 : Form
    {
        i_facility_talEntities db = new i_facility_talEntities();

        public Form1()
        {
            InitializeComponent();
            DayTicker();
            string CorrectedDate = GetCorrectedDate();
            Timer MyTimer = new Timer();
            //MyTimer.Interval = (20 * 1000); // 20 seconds
            MyTimer.Interval = (10 * 1000); // 10 seconds
            MyTimer.Tick += new EventHandler(MyTimer_Tick);
            MyTimer.Start();
        }

        private string GetCorrectedDate()
        {
            string CorrectedDate = "";
            tbldaytiming StartTime1 = db.tbldaytimings.Where(m => m.IsDeleted == 0).FirstOrDefault();
            TimeSpan Start = StartTime1.StartTime;
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


        private void MyTimer_Tick(object sender, EventArgs e)
        {
            string LCorrectedDate, todaysCorrectedDate;

            try
            {
                #region Shift & CorrectedDate
                todaysCorrectedDate = GetCorrectedDate();
                LCorrectedDate = todaysCorrectedDate;//dummy initializaition;

                string Shift = null;
                string CorrectedDate = GetCorrectedDate();
                using (MsqlConnection mcp = new MsqlConnection())
                {
                    mcp.open();
                    String queryshift = "SELECT ShiftName,StartTime,EndTime FROM ["+ MsqlConnection.ServerName + "].[" + MsqlConnection.Schema +"].tblshift_mstr WHERE IsDeleted = 0";
                    SqlDataAdapter dashift = new SqlDataAdapter(queryshift, mcp.sqlConnection);
                    DataTable dtshift = new DataTable();
                    dashift.Fill(dtshift);
                    String[] msgtime = System.DateTime.Now.TimeOfDay.ToString().Split(':');
                    TimeSpan msgstime = System.DateTime.Now.TimeOfDay;
                    //TimeSpan msgstime = new TimeSpan(Convert.ToInt32(msgtime[0]), Convert.ToInt32(msgtime[1]), Convert.ToInt32(msgtime[2]));
                    TimeSpan s1t1 = new TimeSpan(0, 0, 0), s1t2 = new TimeSpan(0, 0, 0), s2t1 = new TimeSpan(0, 0, 0), s2t2 = new TimeSpan(0, 0, 0);
                    TimeSpan s3t1 = new TimeSpan(0, 0, 0), s3t2 = new TimeSpan(0, 0, 0), s3t3 = new TimeSpan(0, 0, 0), s3t4 = new TimeSpan(23, 59, 59);
                    for (int k = 0; k < dtshift.Rows.Count; k++)
                    {
                        if (dtshift.Rows[k][0].ToString().Contains("A"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s1t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s1t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                        else if (dtshift.Rows[k][0].ToString().Contains("B"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s2t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s2t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                        else if (dtshift.Rows[k][0].ToString().Contains("C"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s3t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s3t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                    }

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
                    mcp.close();
                }

                #endregion

                if (System.DateTime.Now.Hour == 6)
                {
                    DayTicker();

                }

                //IsNormalWC = 0 2017-02-10
                //var machineData = db.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsNormalWC == 0).ToList();
                var machineData = db.tblmachinedetails.Where(m => m.IsDeleted == 0).ToList();
                foreach (var macrow in machineData)
                {
                    int machineid = macrow.MachineID;

                    //Insert 1 no_code row when the machine is enabled for the 1st time.
                    #region
                    using (i_facility_talEntities dbquick = new i_facility_talEntities())
                    {
                        var AtLeast1LossRow = dbquick.tbllivelossofentries.Where(m => m.MachineID == machineid).FirstOrDefault();
                        if (AtLeast1LossRow == null)
                        {
                            tbllivelossofentry lossentry = new tbllivelossofentry();
                            lossentry.Shift = Shift.ToString();
                            lossentry.EntryTime = DateTime.Now.AddMinutes(-1);
                            lossentry.StartDateTime = DateTime.Now.AddMinutes(-1);
                            lossentry.EndDateTime = DateTime.Now.AddSeconds(-59);
                            lossentry.CorrectedDate = CorrectedDate;
                            lossentry.IsUpdate = 1;
                            lossentry.DoneWithRow = 1;
                            lossentry.IsStart = 0;
                            lossentry.IsScreen = 0;
                            lossentry.ForRefresh = 0;
                            lossentry.MessageCodeID = 999;
                            int abc = Convert.ToInt32(lossentry.MessageCodeID);
                            var a = dbquick.tbllossescodes.Find(abc);
                            lossentry.MessageDesc = a.LossCodeDesc.ToString();
                            lossentry.MessageCode = a.LossCode.ToString();
                            lossentry.MachineID = machineid;

                            //Session["showIdlePopUp"] = 0;
                            dbquick.tbllivelossofentries.Add(lossentry);
                            dbquick.SaveChanges();
                        }
                    }
                    #endregion
                }

                //Here Call the Stored Procedure
                try
                {

                    #region commented by Ashok
                    //string conString = "server = 'localhost' ;userid = 'root' ;Password = 'srks4$' ;database = 'mazakdaq';port = 3306 ;persist security info=False";
                    //using (MySqlConnection databaseConnection = new MySqlConnection(conString))
                    //{
                    //    MySqlCommand cmd = new MySqlCommand("IdleHandler", databaseConnection);
                    //    cmd.Parameters.AddWithValue("Shift", Shift);
                    //    cmd.Parameters.AddWithValue("CorrectedDate", CorrectedDate);
                    //    databaseConnection.Open();
                    //    cmd.CommandType = CommandType.StoredProcedure;
                    //    var a = cmd.ExecuteNonQuery();
                    //    databaseConnection.Close();
                    //}
                    #endregion
                    idehandler(CorrectedDate, Shift);
                }
                catch (Exception ea)
                {
                    IntoFile(ea.ToString());
                }
                CorrectedDate = GetCorrectedDate();
                //ModeWithLoss(CorrectedDate);
                //ModeVsLossOvelap(CorrectedDate);
                //DeleteExtraLossRows(CorrectedDate);  // Loss Overlap Correction in  livelossofEntry 
            }
            catch (Exception exception)
            {
                IntoFile(exception.ToString());
            }
        }


        private void DayTicker()
        {
            IntoFile("DayTicker:" + DateTime.Now);
            string LCorrectedDate, todaysCorrectedDate;

            try
            {
                #region Shift & CorrectedDate
                todaysCorrectedDate = GetCorrectedDate();
                IntoFile("for todaysCorrectedDate:" + todaysCorrectedDate);
                LCorrectedDate = todaysCorrectedDate;//dummy initializaition;
                IntoFile("for LCorrectedDate:" + LCorrectedDate);
                //correcteddate
                //string correcteddate = null;
                //tbldaytiming StartTime = db.tbldaytimings.Where(m => m.IsDeleted == 0).SingleOrDefault();
                //TimeSpan Start = StartTime.StartTime;
                //if (Start < DateTime.Now.TimeOfDay)
                //{
                //    correcteddate = DateTime.Now.ToString("yyyy-MM-dd");
                //}
                //else
                //{
                //    correcteddate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                //}

                string Shift = null;
                String CorrectedDate = GetCorrectedDate();
                IntoFile("CorrectedDate:" + CorrectedDate);
                using (MsqlConnection mcp = new MsqlConnection())
                {
                    mcp.open();
                    String queryshift = "SELECT ShiftName,StartTime,EndTime FROM ["+ MsqlConnection.ServerName + "].[" + MsqlConnection.Schema +"].tblshift_mstr WHERE IsDeleted = 0";
                    SqlDataAdapter dashift = new SqlDataAdapter(queryshift, mcp.sqlConnection);
                    DataTable dtshift = new DataTable();
                    dashift.Fill(dtshift);
                    String[] msgtime = System.DateTime.Now.TimeOfDay.ToString().Split(':');
                    TimeSpan msgstime = System.DateTime.Now.TimeOfDay;
                    //TimeSpan msgstime = new TimeSpan(Convert.ToInt32(msgtime[0]), Convert.ToInt32(msgtime[1]), Convert.ToInt32(msgtime[2]));
                    TimeSpan s1t1 = new TimeSpan(0, 0, 0), s1t2 = new TimeSpan(0, 0, 0), s2t1 = new TimeSpan(0, 0, 0), s2t2 = new TimeSpan(0, 0, 0);
                    TimeSpan s3t1 = new TimeSpan(0, 0, 0), s3t2 = new TimeSpan(0, 0, 0), s3t3 = new TimeSpan(0, 0, 0), s3t4 = new TimeSpan(23, 59, 59);
                    for (int k = 0; k < dtshift.Rows.Count; k++)
                    {
                        if (dtshift.Rows[k][0].ToString().Contains("A"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s1t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s1t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                        else if (dtshift.Rows[k][0].ToString().Contains("B"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s2t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s2t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                        else if (dtshift.Rows[k][0].ToString().Contains("C"))
                        {
                            String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                            s3t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                            String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                            s3t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                        }
                    }

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
                    mcp.close();
                }

                #endregion

                var machineData = db.tblmachinedetails.Where(m => m.IsDeleted == 0).ToList();
                foreach (var macrow in machineData)
                {
                    int machineid = macrow.MachineID;
                    IntoFile("machineid:" + machineid);
                    //Check for DayChange and replicate BreakdownLoss
                    #region
                    using (i_facility_talEntities dbquick = new i_facility_talEntities())
                    {
                        var LatestBDData = dbquick.tblbreakdowns.Where(m => m.MachineID == machineid).OrderByDescending(m => m.BreakdownID).FirstOrDefault();
                        if (LatestBDData != null)
                        {
                            string tableCorrectedDate = LatestBDData.CorrectedDate;
                            if (LatestBDData.DoneWithRow != 1)
                            {
                                if (CorrectedDate != tableCorrectedDate)
                                {
                                    DateTime TodayStartDate = Convert.ToDateTime(CorrectedDate + " " + new TimeSpan(6, 0, 0));

                                    // Update Endtime for Previous Breakdown loss
                                    LatestBDData.DoneWithRow = 1;
                                    LatestBDData.EndTime = TodayStartDate.AddSeconds(-1);
                                    dbquick.Entry(LatestBDData).State = System.Data.Entity.EntityState.Modified;
                                    dbquick.SaveChanges();

                                    // Insert New Breakdown loss at Day start Time
                                    tblbreakdown tblloe = new tblbreakdown();
                                    tblloe.CorrectedDate = CorrectedDate;
                                    tblloe.DoneWithRow = 0;
                                    tblloe.MachineID = LatestBDData.MachineID;
                                    tblloe.MessageCode = LatestBDData.MessageCode;
                                    tblloe.BreakDownCode = LatestBDData.BreakDownCode;
                                    tblloe.MessageDesc = LatestBDData.MessageDesc;
                                    //tblloe.Shift = "A"; commented by Ashok
                                    tblloe.Shift = Shift;
                                    tblloe.StartTime = TodayStartDate;
                                    dbquick.tblbreakdowns.Add(tblloe);
                                    dbquick.SaveChanges();


                                }
                            }
                        }
                    }
                    #endregion

                    //Check for DayChange and replicate Lossrow
                    #region
                    using (i_facility_talEntities dbquick = new i_facility_talEntities())
                    {
                        var LatestLossData = dbquick.tbllivelossofentries.Where(m => m.MachineID == machineid).OrderByDescending(m => m.LossID).FirstOrDefault();
                        if (LatestLossData != null)
                        {
                            string tableCorrectedDate = LatestLossData.CorrectedDate;
                            if (LatestLossData.DoneWithRow != 1)
                            {
                                if (CorrectedDate != tableCorrectedDate)
                                {
                                    DateTime TodayStartDate = Convert.ToDateTime(CorrectedDate + " " + new TimeSpan(6, 0, 0));

                                    // Update the Endtime of  Previous Day Loss
                                    LatestLossData.IsStart = 0;
                                    LatestLossData.IsScreen = 0;
                                    LatestLossData.ForRefresh = 0;
                                    LatestLossData.DoneWithRow = 1;
                                    LatestLossData.IsUpdate = 1;
                                    LatestLossData.EndDateTime = TodayStartDate.AddSeconds(-1);
                                    dbquick.Entry(LatestLossData).State = System.Data.Entity.EntityState.Modified;
                                    dbquick.SaveChanges();


                                    tbllivelossofentry tblloe = new tbllivelossofentry();
                                    tblloe.CorrectedDate = CorrectedDate;
                                    tblloe.DoneWithRow = 0;
                                    //tblloe.EndDateTime = TodayStartDate;
                                    tblloe.EntryTime = TodayStartDate;
                                    tblloe.ForRefresh = LatestLossData.ForRefresh;
                                    tblloe.IsScreen = LatestLossData.IsScreen;
                                    tblloe.IsStart = LatestLossData.IsStart;
                                    tblloe.IsUpdate = LatestLossData.IsUpdate;
                                    tblloe.MachineID = LatestLossData.MachineID;
                                    tblloe.MessageCode = LatestLossData.MessageCode;
                                    tblloe.MessageCodeID = LatestLossData.MessageCodeID;
                                    tblloe.MessageDesc = LatestLossData.MessageDesc;
                                    tblloe.Shift = Shift;
                                    tblloe.StartDateTime = TodayStartDate;
                                    // Insert New loss at Day start Time
                                    dbquick.tbllivelossofentries.Add(tblloe);
                                    dbquick.SaveChanges();
                                }
                            }
                        }
                    }
                    #endregion

                    //Check for DayChange and replicate Mode 
                    #region
                    using (i_facility_talEntities dbquick = new i_facility_talEntities())
                    {
                        var LatestModeData = dbquick.tbllivemodedbs.Where(m => m.MachineID == machineid && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).FirstOrDefault();
                        if (LatestModeData != null)
                        {
                            string tableCorrectedDate = LatestModeData.CorrectedDate;
                            if (CorrectedDate != tableCorrectedDate)
                            {
                                //DateTime TodayStartDate = Convert.ToDateTime(CorrectedDate + " " + new TimeSpan(6, 0, 0));
                                var daytiming = dbquick.tbldaytimings.Where(m => m.IsDeleted == 0).FirstOrDefault();
                                TimeSpan StartTime = daytiming.StartTime;
                                string Date = CorrectedDate + " " + StartTime;
                                DateTime TodayStartDate = Convert.ToDateTime(Date);

                                //update the end time of Previous Day mode 
                                LatestModeData.IsCompleted = 1;
                                DateTime dt = TodayStartDate.AddSeconds(-1);
                                LatestModeData.EndTime = dt;
                                int duration = (int)Convert.ToDateTime(LatestModeData.EndTime).Subtract(Convert.ToDateTime(LatestModeData.StartTime)).TotalSeconds;
                                LatestModeData.DurationInSec = duration;
                                dbquick.Entry(LatestModeData).State = System.Data.Entity.EntityState.Modified;
                                dbquick.SaveChanges();
                                var LatestModeData1 = dbquick.tbllivemodedbs.Where(m => m.MachineID == machineid && m.IsCompleted == 1).OrderByDescending(m => m.ModeID).FirstOrDefault();

                                string ET = DateTime.Now.Date.ToString("yyyy-MM-dd") + " 05:59:59";
                                DateTime et=Convert.ToDateTime(ET);
                                if (LatestModeData1.EndTime == et)
                                {
                                    tbllivemodedb tblm = new tbllivemodedb();
                                    tblm.CorrectedDate = CorrectedDate;
                                    tblm.ColorCode = LatestModeData.ColorCode;
                                    //tblm.EndTime = TodayStartDate;
                                    tblm.InsertedBy = 1;
                                    tblm.InsertedOn = DateTime.Now;
                                    tblm.IsCompleted = 0;
                                    tblm.IsDeleted = 0;
                                    tblm.MachineID = LatestModeData.MachineID;
                                    tblm.Mode = LatestModeData.Mode;
                                    string ST = DateTime.Now.Date.ToString("yyyy-MM-dd") + " 06:00:00";
                                    tblm.StartTime = Convert.ToDateTime(ST);
                                    // Insert the New mode at day start time
                                    dbquick.tbllivemodedbs.Add(tblm);
                                    dbquick.SaveChanges();
                                }
                                else
                                {
                                    LatestModeData.IsCompleted = 1;
                                    string ET1 = DateTime.Now.Date.ToString("yyyy-MM-dd") + " 05:59:59";
                                    DateTime et1 = Convert.ToDateTime(ET);
                                    LatestModeData.EndTime = et1;
                                    int duration1 = (int)Convert.ToDateTime(LatestModeData.EndTime).Subtract(Convert.ToDateTime(LatestModeData.StartTime)).TotalSeconds;
                                    LatestModeData.DurationInSec = duration1;
                                    dbquick.Entry(LatestModeData).State = System.Data.Entity.EntityState.Modified;
                                    dbquick.SaveChanges();

                                    tbllivemodedb tblm = new tbllivemodedb();
                                    tblm.CorrectedDate = CorrectedDate;
                                    IntoFile("CorrectedDate while inserting:" + tblm.CorrectedDate);
                                    tblm.ColorCode = LatestModeData.ColorCode;
                                    //tblm.EndTime = TodayStartDate;
                                    tblm.InsertedBy = 1;
                                    tblm.InsertedOn = DateTime.Now;
                                    IntoFile("InsertedOn while inserting:" + tblm.InsertedOn);
                                    tblm.IsCompleted = 0;
                                    tblm.IsDeleted = 0;
                                    tblm.MachineID = LatestModeData.MachineID;
                                    tblm.Mode = LatestModeData.Mode;
                                    string ST = DateTime.Now.Date.ToString("yyyy-MM-dd") + " 06:00:00";
                                    tblm.StartTime = Convert.ToDateTime(ST);
                                    IntoFile("StartTime while inserting:" + TodayStartDate);
                                    // Insert the New mode at day start time
                                    dbquick.tbllivemodedbs.Add(tblm);
                                    dbquick.SaveChanges();
                                }
                            }
                        }
                    }
                    #endregion
                }
            }
            catch (Exception exception)
            {
                IntoFile(exception.ToString());
            }
        }

        public void IntoFile(string Msg)
        {
            try
            {
                string path1 = AppDomain.CurrentDomain.BaseDirectory;
                string appPath = Application.StartupPath + @"\LogFileOfIdleHandler.txt";
                using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
                {
                    writer.WriteLine(System.DateTime.Now + ":  " + Msg);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show("IntoFile Error " + e.ToString());
                IntoFile(e.ToString());
            }

        }


        private void idehandler(string correctedDate, string shift)
        {
            IntoFile("idehandler");
            var machineslist = new List<tblmachinedetail>();
            var livelossofentryrow = new tbllivelossofentry();
            var livemoderow = new tbllivemodedb();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                machineslist = db.tblmachinedetails.Where(m => m.IsDeleted == 0).ToList();
            }
            foreach (var machines in machineslist)
            {
                bool isIDLE = false;
                double Totalsec = 0, TotalModesec = 0, durationofCurrent;
                int machineid = machines.MachineID;
                using (i_facility_talEntities db = new i_facility_talEntities())
                {
                    livelossofentryrow = db.tbllivelossofentries.Where(m => m.MachineID == machineid && m.CorrectedDate == correctedDate).OrderByDescending(m => m.LossID).FirstOrDefault();
                    livemoderow = db.tbllivemodedbs.Where(m => m.MachineID == machineid && m.CorrectedDate == correctedDate).OrderByDescending(m => m.ModeID).FirstOrDefault();
                }
                if (livemoderow != null && livelossofentryrow != null)
                {
                    if (livelossofentryrow.IsUpdate == 1 && livelossofentryrow.DoneWithRow == 1)
                    {
                        IntoFile("idehandler:" + livelossofentryrow.IsUpdate);
                        IntoFile("idehandler:" + livelossofentryrow.DoneWithRow);
                        if (livemoderow.Mode == "IDLE")
                        {
                            IntoFile("idehandler:" + livemoderow.Mode);
                            //isIDLE = true;
                            Totalsec = DateTime.Now.Subtract(Convert.ToDateTime(livemoderow.StartTime)).TotalSeconds;
                            TotalModesec = Convert.ToDateTime(livemoderow.InsertedOn).Subtract(Convert.ToDateTime(livemoderow.StartTime)).TotalSeconds;
                            durationofCurrent = Totalsec - TotalModesec;
                            if (durationofCurrent > 120)
                            {
                                IntoFile("idehandler:" + durationofCurrent);
                                tbllivelossofentry tblloe = new tbllivelossofentry();
                                tblloe.MessageCodeID = 999;
                                tblloe.StartDateTime = livemoderow.StartTime;
                                tblloe.EndDateTime = livemoderow.StartTime;
                                tblloe.EntryTime = livemoderow.StartTime;
                                tblloe.CorrectedDate = correctedDate;
                                tblloe.DoneWithRow = 0;
                                tblloe.IsScreen = 0;
                                tblloe.IsStart = 1;
                                tblloe.IsUpdate = 0;
                                tblloe.ForRefresh = 0;
                                tblloe.MachineID = machineid;
                                tblloe.MessageDesc = "No Code Entered";
                                tblloe.Shift = shift;
                                using (i_facility_talEntities db = new i_facility_talEntities())
                                {
                                    db.tbllivelossofentries.Add(tblloe);
                                    db.SaveChanges();
                                }
                            }
                        }
                    }   // Inserting Loss row with  NO code Reason
                    else if (livelossofentryrow.IsUpdate == 0 && livelossofentryrow.DoneWithRow == 0)
                    {
                        IntoFile("idehandler:" + livelossofentryrow.IsUpdate);
                        IntoFile("idehandler:" + livelossofentryrow.DoneWithRow);
                        if (livemoderow.Mode == "IDLE")
                        {
                            isIDLE = true;
                        }
                        else
                        {
                            IntoFile("idehandler:" + livemoderow.Mode);
                            var livelossofentry = new tbllivelossofentry();
                            using (i_facility_talEntities db = new i_facility_talEntities())
                            {
                                livelossofentry = db.tbllivelossofentries.Find(livelossofentryrow.LossID);
                            }
                            UpdateLossofEntries(livelossofentry.LossID, Convert.ToDateTime(livemoderow.StartTime));
                        }
                    }    // To update the Loss Reasonnew code at 1st Time
                    else if (livelossofentryrow.IsUpdate == 1 && livelossofentryrow.DoneWithRow == 0)
                    {
                        IntoFile("idehandler:" + livelossofentryrow.IsUpdate);
                        IntoFile("idehandler:" + livelossofentryrow.DoneWithRow);
                        Totalsec = DateTime.Now.Subtract(Convert.ToDateTime(livelossofentryrow.EntryTime)).TotalSeconds;
                        if (livemoderow.Mode == "IDLE")
                        {
                            isIDLE = true;
                        }
                        else if (!isIDLE)
                        {
                            IntoFile("idehandler:" + livemoderow.Mode);
                            var livelossofentry = new tbllivelossofentry();
                            using (i_facility_talEntities db = new i_facility_talEntities())
                            {
                                livelossofentry = db.tbllivelossofentries.Find(livelossofentryrow.LossID);
                            }
                            UpdateLossofEntries(livelossofentry.LossID, Convert.ToDateTime(livemoderow.StartTime));
                        }
                        else if (Totalsec > 120)
                        {
                            var livelossofentry = new tbllivelossofentry();
                            using (i_facility_talEntities db = new i_facility_talEntities())
                            {
                                livelossofentry = db.tbllivelossofentries.Find(livelossofentryrow.LossID);
                            }
                            if (livelossofentry != null)
                            {
                                livelossofentry.IsScreen = 1;
                                using (i_facility_talEntities db1 = new i_facility_talEntities())
                                {
                                    db1.Entry(livelossofentry).State = System.Data.Entity.EntityState.Modified;
                                    db1.SaveChanges();
                                }
                            }
                        }
                    }    //  To update the new code at 2nd Time and so on.  
                }
                else if (livelossofentryrow == null && livemoderow != null)
                {
                    if (livemoderow.Mode == "IDLE")
                    {
                        //isIDLE = true;
                        Totalsec = DateTime.Now.Subtract(Convert.ToDateTime(livemoderow.StartTime)).TotalSeconds;
                        TotalModesec = Convert.ToDateTime(livemoderow.InsertedOn).Subtract(Convert.ToDateTime(livemoderow.StartTime)).TotalSeconds;
                        durationofCurrent = Totalsec - TotalModesec;
                        if (durationofCurrent > 120)
                        {
                            tbllivelossofentry tblloe = new tbllivelossofentry();
                            tblloe.MessageCodeID = 999;
                            tblloe.StartDateTime = livemoderow.StartTime;
                            tblloe.EndDateTime = livemoderow.StartTime;
                            tblloe.EntryTime = livemoderow.StartTime;
                            tblloe.CorrectedDate = correctedDate;
                            tblloe.DoneWithRow = 0;
                            tblloe.IsScreen = 0;
                            tblloe.IsStart = 1;
                            tblloe.IsUpdate = 0;
                            tblloe.ForRefresh = 0;
                            tblloe.MachineID = machineid;
                            tblloe.MessageDesc = "No Code Entered";
                            tblloe.Shift = shift;
                            using (i_facility_talEntities db = new i_facility_talEntities())
                            {
                                db.tbllivelossofentries.Add(tblloe);
                                db.SaveChanges();
                            }
                        }
                    }
                }
            }

        }

        private void UpdateLossofEntries(int LossID, DateTime ModeStartTime)
        {
            IntoFile("UpdateLossofEntries:" + ModeStartTime);
            var livelossofentry = new tbllivelossofentry();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                livelossofentry = db.tbllivelossofentries.Find(LossID);
            }

            if (livelossofentry.StartDateTime <= ModeStartTime)
            {
                livelossofentry.EndDateTime = ModeStartTime;
                livelossofentry.DoneWithRow = 1;
                livelossofentry.IsUpdate = 1;
                livelossofentry.IsScreen = 0;
                livelossofentry.IsStart = 0;
                livelossofentry.ForRefresh = 0;
                using (i_facility_talEntities db1 = new i_facility_talEntities())
                {
                    db1.Entry(livelossofentry).State = System.Data.Entity.EntityState.Modified;
                    db1.SaveChanges();
                }
            }

        }
    }
}
