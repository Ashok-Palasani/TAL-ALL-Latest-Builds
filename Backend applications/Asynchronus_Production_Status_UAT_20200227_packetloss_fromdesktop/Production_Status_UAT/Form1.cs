﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Production_Status_UAT
{
    public partial class Production_Status_UAT : Form
    {
        private i_facility_talEntities db = new i_facility_talEntities();
        //int Fourthgencount = 0;
        public Production_Status_UAT()
        {
            InitializeComponent();


            Method_ForDeleting();
            //string CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
            //GetCorrecteddate(CorrectedDate);
            try
            {
                MachineList();
            }
            catch (Exception ex)
            {
                IntoFile(ex.ToString());
            }
            Timer MyTimer = new Timer();
            MyTimer.Interval = (45 * 1000); // 45 Sec
            MyTimer.Tick += new EventHandler(MyTimer_Tick);
            MyTimer.Start();
            

            try
            {
                Timer MyTimerDelete = new Timer();
                MyTimerDelete.Interval = (5 * 60 * 1000 * 12); // 1 Hr
                MyTimerDelete.Tick += new EventHandler(MyTimer_ForDeletingFromPCBDAQINTable);
                MyTimerDelete.Start();
            }

            catch (Exception e)
            {
                IntoFile(" Main Catch " + e.ToString());
            }
        }


        private void MyTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                MachineList();
            }
            catch (Exception ex)
            {
                IntoFile(ex.ToString());
            }
        }


        public async Task MachineList()
        {

            #region SHIFT and CorrectedDate

            string Shift = null;
            DataTable dtshift = new DataTable();
            String queryshift = "SELECT ShiftName,StartTime,EndTime FROM [i_facility_tsal].[dbo].shift_master WHERE IsDeleted = 0 ";
            ConnectionString mcp = new ConnectionString();
            mcp.open();
            using (SqlDataAdapter dashift = new SqlDataAdapter(queryshift, mcp.msqlConnection))
            {
                dashift.Fill(dtshift);
            }
            mcp.close();

            String[] msgtime = System.DateTime.Now.TimeOfDay.ToString().Split(':');
            TimeSpan msgstime = System.DateTime.Now.TimeOfDay;
            //TimeSpan msgstime = new TimeSpan(Convert.ToInt32(msgtime[0]), Convert.ToInt32(msgtime[1]), Convert.ToInt32(msgtime[2]));
            TimeSpan s1t1 = new TimeSpan(0, 0, 0), s1t2 = new TimeSpan(0, 0, 0), s2t1 = new TimeSpan(0, 0, 0), s2t2 = new TimeSpan(0, 0, 0);
            TimeSpan s3t1 = new TimeSpan(0, 0, 0), s3t2 = new TimeSpan(0, 0, 0), s3t3 = new TimeSpan(0, 0, 0), s3t4 = new TimeSpan(23, 59, 59);
            for (int k = 0; k < dtshift.Rows.Count; k++)
            {
                if (dtshift.Rows[k][0].ToString().Contains("1"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s1t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                    s1t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                }
                else if (dtshift.Rows[k][0].ToString().Contains("2"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s2t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
                    s2t2 = new TimeSpan(Convert.ToInt32(s11[0]), Convert.ToInt32(s11[1]), Convert.ToInt32(s11[2]));
                }
                else if (dtshift.Rows[k][0].ToString().Contains("3"))
                {
                    String[] s1 = dtshift.Rows[k][1].ToString().Split(':');
                    s3t1 = new TimeSpan(Convert.ToInt32(s1[0]), Convert.ToInt32(s1[1]), Convert.ToInt32(s1[2]));
                    String[] s11 = dtshift.Rows[k][2].ToString().Split(':');
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

            DataTable dataHolderj = new DataTable();
            try
            {
                ConnectionString mc = new ConnectionString();
                mc.open();
                //String sql1 = "Select * from ( select DAQINID ,v.MachineID , d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn from [i_facility_tsal].[dbo].pcbdaqin_tblNew as d , VW_Join_DetailsANDParam as v "
                //                + " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc ) as T "
                //                + " group by T.PCBIPAddress,T.ParamPIN";

                string sql1 = "Select d.DAQINID ,v.MachineID , d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn from [i_facility_tsal].[dbo].pcbdaqin_tblNew as d , vw_join_detailsandparamNew as v "
                                + " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc";

                using (SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection))
                {
                    da1.Fill(dataHolderj);
                }



                DataView dt1 = new DataView(dataHolderj);
                dt1.Sort = "PCBIPAddress ASC,ParamPIN ASC";
                DataTable dtnew = dt1.ToTable();

                IEnumerable<DataRow> qryLatestInterview = from rows in dtnew.AsEnumerable()
                                                          group rows by new { PositionID = rows["PCBIPAddress"], CandidateID = rows["ParamPIN"] } into grp
                                                          select grp.First();
                dtnew = qryLatestInterview.CopyToDataTable();
                dataHolderj = dtnew;
                mc.close();
            }
            catch (Exception e)
            {
                IntoFile(" pcbdaqin_tblNew " + e.ToString());
            }
            List<tblmachinedetail> machineList = new List<tblmachinedetail>();
            using (i_facility_talEntities db1 = new i_facility_talEntities())
            {
                #region For NON-HAAS Machines => IsLevel != 4
                //var machineList = db1.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsPCB == 1 && m.MachineID == 1); //&& m.MachineID == 24
                machineList = db1.tblmachinedetails.Where(m => m.IsDeleted == 0 && m.IsPCB == 1).OrderBy(m=>m.MachineID).ToList(); //&& m.MachineID == 24
            }

          //  await DAQ(machineList, dataHolderj, CorrectedDate, Shift);
            foreach (tblmachinedetail machine in machineList)
            {
                int machineId = machine.MachineID;
                int isfourthGen = machine.IsDLVersion;
                //dont push if last mode is Breakdown
                int IsBreakdown = 0;
                //IntoFile("MachineID:" + machineId + " MachineIPAddress :" + machine.IPAddress);
                // var lastMode = new tbllivemodedb();
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    //lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).OrderByDescending(m => m.InsertedOn).Take(1).SingleOrDefault();
                    tbllivemodedb lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).FirstOrDefault();
                    if (lastMode != null)
                    {
                        string mode = lastMode.Mode;
                        int modeStatus = lastMode.IsCompleted;
                        if (mode == "BREAKDOWN")
                        {
                            IsBreakdown = 1;
                        }
                        if (mode == "MNT")
                        {
                            IsBreakdown = 1;
                        }
                    }
                    else
                    {
                        IsBreakdown = 0;
                    }
                }


                if (IsBreakdown == 0)
                {
                    string pcbipaddress = "";
                    List<pcb_parameters> ParamList = new List<pcb_parameters>();
                    using (i_facility_talEntities db5 = new i_facility_talEntities())
                    {
                        //string pcbipaddress = db.pcb_details.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).Select(m => m.PCBIPAddress).FirstOrDefault();
                        pcbipaddress = db5.pcb_details.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).Select(m => m.PCBIPAddress).FirstOrDefault();
                        ParamList = db5.pcb_parameters.Where(m => m.MachineID == machineId && m.IsDeleted == 0 && m.ParameterType == "PON").ToList();

                        //IntoFile("MachineID:" + machineId + " PCBIPAddress :" + pcbipaddress);

                    }
                    int count = 0;

                    TryPing:
                    bool pingable = false;
                    Ping pinger = new Ping();
                    try
                    {

                        PingReply reply = pinger.Send(pcbipaddress, 2000); //2000
                        pingable = reply.Status == IPStatus.Success;
                        //IntoFile1("before fourth gen Pinging Status of " + pcbipaddress + ": " + pingable);

                        // Removed the commented
                        if (isfourthGen == 1)
                        {
                            IntoFile1("Pinging Status of " + pcbipaddress + ": " + pingable);
                            pingable = true;
                        }

                        IntoFile("Pinging Status of " + pcbipaddress + ": " + pingable);

                        if (!pingable)
                        {
                            IntoFile1("Pinging Status of " + pcbipaddress + ": " + pingable);
                        }
                    }
                    catch (PingException e)
                    {
                        // Discard PingExceptions and return false;
                        IntoFile(" PING Exception" + "MachineID: " + machineId + " " + e.ToString());
                    }
                    catch (Exception e)
                    {
                        IntoFile(" During Ping " + "MachineID: " + machineId + " " + e.ToString());
                    }


                    if (pingable)
                    {
                        #region OLD NO-PING NO-BLUE
                        handle_no_ping NoPingData = new handle_no_ping();
                        using (i_facility_talEntities db5 = new i_facility_talEntities())
                        {
                            NoPingData = db5.handle_no_ping.Where(m => m.MachineID == machineId).FirstOrDefault();
                        }
                        if (NoPingData != null)
                        {
                            int id = NoPingData.NoPingID;
                            handle_no_ping hnp = db.handle_no_ping.Find(id);
                            db.handle_no_ping.Remove(hnp);
                            db.SaveChanges();
                        }
                        #endregion
                        bool ispacketloss = false;
                        IOTGatwayPacketsData IOTgatewaydata = db.IOTGatwayPacketsDatas.Where(m => m.IPAddres == pcbipaddress).OrderByDescending(m => m.GatewayMsgID).FirstOrDefault();

                        if (IOTgatewaydata != null)
                        {
                            DateTime DBdatetime = Convert.ToDateTime(IOTgatewaydata.CreatedOn);
                            DateTime CurrentTime = DateTime.Now;
                            double diff = CurrentTime.Subtract(DBdatetime).TotalMinutes;
                            if (diff >= 3)
                            {
                                ispacketloss = true;
                            }
                        }

                        if (!ispacketloss)
                        {
                            //Delete the Machine No Ping details from the DB.
                            PINGHANDLER_tbl dldata1 = db.PINGHANDLER_tbl.Where(m => m.MachineID == machineId && m.IPADDRESS == pcbipaddress).OrderByDescending(m => m.PID).FirstOrDefault();
                            if (dldata1 != null)
                            {
                                db.PINGHANDLER_tbl.Remove(dldata1);
                            }
                            List<DataRow> dr = dataHolderj.Select("MachineID = '" + @machineId + "'").ToList();
                            if (dr != null)
                            {
                                //Set Default values to All Pins
                                int pin17 = 0, pin18 = 0, pin19 = 0, pin20 = 0, pin22 = 0, pin23 = 0, pin24 = 0, pin25 = 0, pin26 = 0;
                                List<KeyValuePair<int, int>> Factorlist = new List<KeyValuePair<int, int>>();
                                //list.Add(new KeyValuePair<string, int>("Rabbit", 4));
                                #region Try Catch 1
                                try
                                {
                                    foreach (DataRow rowj in dr)
                                    {
                                        //var jack = rowj;
                                        int pinNo = Convert.ToInt32(rowj[3]);
                                        int pinVal = Convert.ToInt32(rowj[4]);
                                        IEnumerable<pcb_parameters> paramData = (from row in ParamList where row.PinNumber == pinNo select row);
                                        if (paramData != null)
                                        {
                                            Factorlist.Add(new KeyValuePair<int, int>(pinNo, pinVal));
                                        }

                                        //Now Assign Value taken from pcbdaqin_tblNew to PIN's
                                        switch (pinNo)
                                        {
                                            case 17:
                                                {
                                                    pin17 = pinVal;
                                                    break;
                                                }
                                            case 18:
                                                {
                                                    pin18 = pinVal;
                                                    break;
                                                }
                                            case 19:
                                                {
                                                    pin19 = pinVal;
                                                    break;
                                                }
                                            case 20:
                                                {
                                                    pin20 = pinVal;
                                                    break;
                                                }
                                            case 22:
                                                {
                                                    pin22 = pinVal;
                                                    break;
                                                }
                                            case 23:
                                                {
                                                    pin23 = pinVal;
                                                    break;
                                                }
                                            case 24:
                                                {
                                                    pin24 = pinVal;
                                                    break;
                                                }
                                            case 25:
                                                {
                                                    pin25 = pinVal;
                                                    break;
                                                }
                                            case 26:
                                                {
                                                    pin26 = pinVal;
                                                    break;
                                                }
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    IntoFile(" Gen KeyValuePair " + "MachineID: " + machineId + " " + e.ToString());
                                    //MessageBox.Show("Gen KeyValuePair " + e);
                                }

                                #endregion //try catch 1

                                #region Try Catch 2
                                try
                                {
                                    //int PONHighValue = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.HighValue).SingleOrDefault());
                                    //int PONPinNo = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.PinNumber).SingleOrDefault());
                                    int PONHighValue = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.HighValue).FirstOrDefault());
                                    int PONPinNo = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.PinNumber).FirstOrDefault());

                                    List<int> ponvalue = (from kvp in Factorlist where kvp.Key == PONPinNo select kvp.Value).ToList();

                                    if (ponvalue.Count > 0)
                                    {
                                        int PONValue = ponvalue[0];

                                        if (PONHighValue == PONValue)
                                        {
                                            #region //New Logic to Decide the Color( Row ) into  tbllivemodedb Table.
                                            string ColorIntoTblMode = "yellow";
                                            pcbdps_master ColorData = new pcbdps_master();
                                            using (i_facility_talEntities db5 = new i_facility_talEntities())
                                            {
                                                ColorData = db.pcbdps_master.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.Pin17 == pin17 && m.Pin18 == pin18 && m.Pin19 == pin19 && m.Pin20 == pin20 && m.Pin22 == pin22 && m.Pin23 == pin23 && m.Pin24 == pin24 && m.Pin25 == pin25 && m.Pin26 == pin26).FirstOrDefault();
                                            }
                                            if (ColorData != null)
                                            {
                                                ColorIntoTblMode = ColorData.ColorValue;
                                            }
                                            if (ColorIntoTblMode == "yellow")
                                            {
                                                await InsertingDataIntoModeTable("IDLE", machineId, ColorIntoTblMode, CorrectedDate, Shift);
                                            }
                                            else if (ColorIntoTblMode == "green")
                                            {
                                                await InsertingDataIntoModeTable("PowerOn", machineId, ColorIntoTblMode, CorrectedDate, Shift);
                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            #region Check in tbllivemodedb and push. TODO

                                            tbllivemodedb lastMode = new tbllivemodedb();

                                            using (i_facility_talEntities db5 = new i_facility_talEntities())
                                            {
                                                lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.InsertedOn).FirstOrDefault();
                                            }
                                            if (lastMode != null)
                                            {
                                                string mode = lastMode.Mode;
                                                int modeStatus = lastMode.IsCompleted;
                                                if (mode == "BREAKDOWN")
                                                {
                                                    //await InsertingDataIntoModeTable("BREAKDOWN", machineId, "red", CorrectedDate, Shift);
                                                }
                                                else if (mode == "MNT")
                                                {
                                                    //await InsertingDataIntoModeTable("MNT", machineId, "red", CorrectedDate, Shift);
                                                }
                                                else
                                                {
                                                    await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                                                }
                                            }
                                            else
                                            {
                                                await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    IntoFile(" After KeyGen " + "MachineID: " + machineId + " " + e.ToString());
                                    //MessageBox.Show("After KeyGen: " + e);
                                }
                                #endregion  //end of try catch 2
                            }
                        }
                        else
                        {
                            await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                        }
                    }
                    else //if not pingable
                    {
                        count++;
                        //Get Db Counter for No Pingable 4th Gen DL
                        //PINGHANDLER_tbl dldata = new PINGHANDLER_tbl();
                        //int DlCounter = 0;
                        //using (i_facility_talEntities db = new i_facility_talEntities())
                        //{
                        //    dldata = db.PINGHANDLER_tbl.Where(m => m.MachineID == machineId).OrderByDescending(m => m.PID).FirstOrDefault();
                        //}
                        //if (dldata != null)
                        //{
                        //    DlCounter = Convert.ToInt32(dldata.Pingcount);
                        //}


                        if (count < 2) //5
                        {
                            goto TryPing;
                        }
                        //else if (isfourthGen == 1 && DlCounter <= 6)
                        //{
                        //    DlCounter += 1;
                        //    //Insert or update into DB
                        //    if (dldata != null)
                        //    {
                        //        dldata.MachineID = machineId;
                        //        dldata.IPADDRESS = pcbipaddress;
                        //        dldata.CorrectedDate = CorrectedDate;
                        //        dldata.Pingcount = 1;
                        //        try
                        //        {
                        //            using (i_facility_talEntities db5 = new i_facility_talEntities())
                        //            {
                        //                db5.PINGHANDLER_tbl.Add(dldata);
                        //                db5.SaveChanges();
                        //            }
                        //        }
                        //        catch (Exception)
                        //        {
                        //            //IntoFile("AddModeRow Method For MachineID:" + machineId + "Error is" + e.ToString());
                        //        }
                        //    }
                        //    else
                        //    {
                        //        dldata.Pingcount += 1;
                        //        using (i_facility_talEntities db5 = new i_facility_talEntities())
                        //        {
                        //            db5.Entry(dldata).State = System.Data.Entity.EntityState.Modified;
                        //            db5.SaveChanges();
                        //        }
                        //    }

                        //    goto TryPing;    // Mode will delay's due to this
                        //}

                        else
                        {
                            //if (isfourthGen == 1)
                            //{
                            //    //pingable = true;
                            //}
                            //else
                            await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                        }
                    }
                }
            }
            #endregion

        }


        public async Task<bool> InsertingDataIntoModeTable(string mode, int machineId, string color, string CorrectedDate, string Shift)
        {
            bool tic = true;
            //Based on the machineid and correctedDate select the last(Latest) mode in mode table 
            // see if new mode is equal to last(latest) mode , if not insert.

            //Take last mode from today insert it now. or latest of modes from latest of previous days . Only if current mode is different from new mode.

            //var lastMode = db.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.CorrectedDate == CorrectedDate).OrderByDescending(m => m.InsertedOn).Take(1).SingleOrDefault();

            #region cheking the day for completion

            List<tbllivemodedb> lastMode = new List<tbllivemodedb>();
            tbllivemodedb lastmodesinglerow = new tbllivemodedb();

            using (i_facility_talEntities db5 = new i_facility_talEntities())
            {
                //lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.CorrectedDate == CorrectedDate).OrderByDescending(m => m.ModeID).FirstOrDefault();
                //lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.CorrectedDate == CorrectedDate && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).ToList();
                lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).ToList();
            }

            if (lastMode != null && lastMode.Count == 1)
            {
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    //lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.CorrectedDate == CorrectedDate).OrderByDescending(m => m.ModeID).FirstOrDefault();
                    //lastmodesinglerow = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.CorrectedDate == CorrectedDate && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).FirstOrDefault();
                    lastmodesinglerow = lastMode.OrderByDescending(m => m.ModeID).FirstOrDefault();
                }


                if (lastmodesinglerow.Mode != mode)
                {
                    try
                    {
                        //update endtime for last mode 
                        DateTime dt = DateTime.Now;
                        int lastmodeID = lastmodesinglerow.ModeID;

                        //get colorcode
                        //string color = null;
                        if (lastmodeID != 0)
                        {
                            tbllivemodedb tmprevious = new tbllivemodedb();

                            //v changes

                            DateTime CompletedModeET = DateTime.Now;
                            using (i_facility_talEntities db5 = new i_facility_talEntities())
                            {
                                // getting the last completed endtime validete with the present start time
                                CompletedModeET = Convert.ToDateTime(db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.IsCompleted == 1 && m.MachineID == machineId).OrderByDescending(m => m.ModeID).Select(m => m.EndTime).FirstOrDefault());
                                tmprevious = db.tbllivemodedbs.Find(lastmodeID);

                            }
                            DateTime PresentModeST = Convert.ToDateTime(tmprevious.StartTime);

                            if (PresentModeST == CompletedModeET)
                            {
                                string previousmode = tmprevious.Mode;
                                if (mode != previousmode)
                                {
                                    UpdateModeRow(dt, lastmodeID, machineId, color, CorrectedDate).Wait();
                                    //UpdateModeRow(dt, lastmodeID, machineId).Wait();
                                    #region Commented By Ashok
                                    //if (mode == "PowerOff")
                                    //{
                                    //    color = "blue";
                                    //}
                                    //else if (mode == "PowerOn")
                                    //{
                                    //    color = "green";
                                    //}
                                    //else if (mode == "IDLE")
                                    //{
                                    //    color = "yellow";
                                    //}
                                    #endregion
                                    AddModeRow(CorrectedDate, dt, color, machineId, mode).Wait();
                                }
                            }
                            else
                            {
                                string previousmode = tmprevious.Mode;
                                if (mode != previousmode)
                                {
                                    //UpdateModeRowWithStartTime(dt, lastmodeID, machineId, CompletedModeET, color, CorrectedDate).Wait();
                                    //UpdateModeRowWithStartTime(dt, lastmodeID, machineId, CompletedModeET).Wait();
                                    UpdateModeRow(dt, lastmodeID, machineId, color, CorrectedDate).Wait();
                                    #region Commented by Ashok
                                    //if (mode == "PowerOff")
                                    //{
                                    //    color = "blue";
                                    //}
                                    //else if (mode == "PowerOn")
                                    //{
                                    //    color = "green";
                                    //}
                                    //else if (mode == "IDLE")
                                    //{
                                    //    color = "yellow";
                                    //}
                                    #endregion
                                    AddModeRow(CorrectedDate, dt, color, machineId, mode).Wait();
                                }
                            }

                        }
                    }
                    catch (Exception e)
                    {
                        IntoFile("In Mode ::  1st Loop " + "MachineID: " + machineId + " " + e.ToString());
                        //MessageBox.Show("In Mode ::  1st Loop " + e);
                    }
                }
            }
            else if (lastMode.Count > 1)
            {
                string getmodequery2 = "SELECT ModeID,StartTime,Mode From [i_facility_tsal].[dbo].tbllivemodedb WHERE IsCompleted = 0 and MachineID = " + machineId + " and CorrectedDate<='" + CorrectedDate + "' order by ModeID";
                DataTable dtModeMultiple = new DataTable();

                using (ConnectionString mc = new ConnectionString())
                {
                    mc.open();
                    SqlDataAdapter daMode1 = new SqlDataAdapter(getmodequery2, mc.msqlConnection);
                    daMode1.Fill(dtModeMultiple);
                    mc.close();
                }
                using (ConnectionString mc = new ConnectionString())
                {
                    for (int i = 0; i < (dtModeMultiple.Rows.Count - 1); i++)
                    {
                        if (dtModeMultiple.Rows[i][2].ToString() == dtModeMultiple.Rows[i + 1][2].ToString())
                        {
                            mc.open();
                            SqlCommand cmdpoweroff = new SqlCommand("DELETE FROM [i_facility_tsal].[dbo].tbllivemodedb where ModeID = " + Convert.ToInt32(dtModeMultiple.Rows[i][0]), mc.msqlConnection);
                            int ret = cmdpoweroff.ExecuteNonQuery();
                            mc.close();
                            IntoFile("DeletedPreviousUncompletedModes ," + dtModeMultiple.Rows[i][0].ToString());
                        }
                    }
                }
            }
            else
            {
                //List<tbllivemodedb> previousMode = new List<tbllivemodedb>();
                //tbllivemodedb previousModesinglerow = new tbllivemodedb();
                //using (i_facility_talEntities db5 = new i_facility_talEntities())
                //{
                //    //previousMode = db.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).OrderByDescending(m => m.ModeID).FirstOrDefault();
                //    previousMode = db.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).ToList();
                //}
                #region Commented By Ashok
                //if (previousMode != null && previousMode.Count == 1)
                //{
                //    previousModesinglerow = db.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).FirstOrDefault();
                //    // v condition
                //    try
                //    {
                //        DateTime dt = DateTime.Now;
                //        int id = previousModesinglerow.ModeID;

                //        //get colorcode
                //        //string color = null;
                //        if (id != 0)
                //        {
                //            tbllivemodedb tmprevious = new tbllivemodedb();

                //            DateTime CompletedModeET = DateTime.Now;
                //            using (i_facility_talEntities db5 = new i_facility_talEntities())
                //            {
                //                CompletedModeET = Convert.ToDateTime(db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.IsCompleted == 1 && m.MachineID == machineId).OrderByDescending(m => m.ModeID).Select(m => m.EndTime).FirstOrDefault());
                //                tmprevious = db.tbllivemodedbs.Find(id);
                //            }
                //            DateTime PresentModeST = Convert.ToDateTime(tmprevious.StartTime);

                //            if (PresentModeST == CompletedModeET)
                //            {
                //                string previousmode = tmprevious.Mode;
                //                //if (mode != previousmode)
                //                //{
                //                //tmprevious.EndTime = dt;
                //                //tmprevious.IsCompleted = 1;
                //                //try
                //                //{
                //                //    using (i_facility_talEntities db5 = new i_facility_talEntities())
                //                //    {
                //                //        db5.Entry(tmprevious).State = System.Data.Entity.EntityState.Modified;
                //                //        db5.SaveChanges();
                //                //    }
                //                //}
                //                //catch (Exception e)
                //                //{
                //                //    IntoFile("MachineID: " + machineId + ". InsetingDataIntoModeTable ::  " + e.ToString());
                //                //}

                //                UpdateModeRow(dt, id, machineId, color, CorrectedDate).Wait();
                //                //UpdateModeRow(dt, id, machineId).Wait();


                //                //tbllivemodedb tm = new tbllivemodedb();
                //                //tm.CorrectedDate = CorrectedDate;
                //                //tm.InsertedBy = 1;
                //                //tm.InsertedOn = dt;
                //                //tm.StartTime = dt;
                //                //tm.ColorCode = previousMode.ColorCode;
                //                //tm.IsDeleted = 0;
                //                //tm.IsCompleted = 0;
                //                //tm.MachineID = machineId;
                //                //tm.Mode = previousMode.Mode;
                //                //try
                //                //{
                //                //    using (i_facility_talEntities db5 = new i_facility_talEntities())
                //                //    {
                //                //        db5.tbllivemodedbs.Add(tm);
                //                //        db5.SaveChanges();
                //                //    }
                //                //}
                //                //catch (Exception e)
                //                //{
                //                //    IntoFile("MachineID: " + machineId + ". InsertingDataIntoModeTable 921 ::  " + e.ToString());
                //                //}
                //                //}


                //                AddModeRow(CorrectedDate, dt, previousModesinglerow.ColorCode, machineId, previousModesinglerow.Mode).Wait();

                //            }
                //            else
                //            {
                //                string previousmode = tmprevious.Mode;
                //                //if (mode != previousmode)
                //                //{
                //                //using (i_facility_talEntities db5 = new i_facility_talEntities())

                //                //{
                //                //    tbllivemodedb tmlastmode = db5.tbllivemodedbs.Find(id);
                //                //    tmlastmode.StartTime = CompletedModeET;
                //                //    tmlastmode.EndTime = dt;
                //                //    tmlastmode.IsCompleted = 1;
                //                //    db5.Entry(tmlastmode).State = System.Data.Entity.EntityState.Modified;
                //                //    db5.SaveChanges();
                //                //}
                //                IntoFile("dt, id, machineId, CompletedModeET" + dt + "," + id + "," + machineId + "," + CompletedModeET);
                //                UpdateModeRowWithStartTime(dt, id, machineId, CompletedModeET,color,CorrectedDate).Wait();
                //                //UpdateModeRowWithStartTime(dt, id, machineId, CompletedModeET).Wait();
                //                if (mode == "PowerOff")
                //                {
                //                    color = "blue";
                //                }
                //                else if (mode == "PowerOn")
                //                {
                //                    color = "green";
                //                }
                //                else if (mode == "IDLE")
                //                {
                //                    color = "yellow";
                //                }

                //                //tbllivemodedb tm = new tbllivemodedb();
                //                //tm.CorrectedDate = CorrectedDate;
                //                //tm.InsertedBy = 1;
                //                //tm.InsertedOn = dt;
                //                //tm.StartTime = dt;
                //                //tm.ColorCode = color;
                //                //tm.IsDeleted = 0;
                //                //tm.IsCompleted = 0;
                //                //tm.MachineID = machineId;
                //                //tm.Mode = mode;
                //                //using (i_facility_talEntities db5 = new i_facility_talEntities())
                //                //{
                //                //    db5.tbllivemodedbs.Add(tm);
                //                //    db5.SaveChanges();
                //                //}
                //                //}

                //                AddModeRow(CorrectedDate, dt, color, machineId, mode).Wait();
                //            }
                //        }
                //    }
                //    catch (Exception e)
                //    {
                //        IntoFile("MachineID: " + machineId + ". In Mode ::  2nd Loop " + e.ToString());
                //        //MessageBox.Show("In Mode ::  2st Loop " + e);
                //    }

                //}
                //else if (previousMode.Count > 1)
                //{
                //    try
                //    {
                //        DataTable dtMode1 = new DataTable();
                //        String getmodequery1 = "SELECT ModeID,StartTime From [i_facility_tsal].[dbo].tbllivemodedb WHERE IsCompleted = 0 and MachineID = " + machineId + " order by ModeID";
                //        using (ConnectionString mc = new ConnectionString())
                //        {
                //            mc.open();
                //            SqlDataAdapter daMode1 = new SqlDataAdapter(getmodequery1, mc.msqlConnection);
                //            daMode1.Fill(dtMode1);
                //            mc.close();
                //        }
                //        using (ConnectionString mc = new ConnectionString())
                //        {
                //            for (int i = 0; i < (dtMode1.Rows.Count - 2); i++)
                //            {
                //                if (dtMode1.Rows[i][1].ToString() == DateTime.Now.ToString("yyyy-MM-dd 07:15:00"))
                //                {
                //                    mc.open();
                //                    SqlCommand cmdpoweroff = new SqlCommand("DELETE FROM [i_facility_tsal].[dbo].tbllivemodedb where ModeID = " + Convert.ToInt32(dtMode1.Rows[i][0]), mc.msqlConnection);
                //                    int ret = cmdpoweroff.ExecuteNonQuery();
                //                    mc.close();

                //                }
                //            }
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        IntoFile(ex.ToString());
                //    }
                //}

                #endregion
                /* else*/ // No rows in tbllivemodedb for this machine.
                {
                    DateTime dt = DateTime.Now;
                    AddModeRow(CorrectedDate, dt, color, machineId, mode).Wait();

                }
            }
            #endregion

            return await Task<int>.FromResult(tic);
        }

        public async Task<int> AddModeRow(string CorrectedDate, DateTime dt, string color, int machineId, string mode)
        {
            int Ret = 1;
            tbllivemodedb tm = new tbllivemodedb();
            tm.CorrectedDate = CorrectedDate;
            tm.InsertedBy = 1;
            tm.InsertedOn = dt;
            tm.StartTime = dt;
            tm.ColorCode = color;
            tm.IsDeleted = 0;
            tm.IsCompleted = 0;
            tm.MachineID = machineId;
            tm.Mode = mode;
            try
            {
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    db5.tbllivemodedbs.Add(tm);
                    db5.SaveChanges();
                }
            }
            catch (Exception e)
            {
                IntoFile("AddModeRow Method For MachineID:" + machineId + "Error is" + e.ToString());
            }

            return await Task<int>.FromResult(Ret);
        }

        public async Task<int> UpdateModeRow(DateTime dt, int lastmachineid, int machineId)
        {
            int Ret = 1;
            tbllivemodedb tmprevious = new tbllivemodedb();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                tmprevious = db.tbllivemodedbs.Find(lastmachineid);
            }
            tmprevious.EndTime = dt;
            tmprevious.IsCompleted = 1;
            double duationinsec = dt.Subtract(Convert.ToDateTime(tmprevious.StartTime)).TotalSeconds;
            tmprevious.DurationInSec = Convert.ToInt32(duationinsec);
            try
            {
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    db5.Entry(tmprevious).State = System.Data.Entity.EntityState.Modified;
                    db5.SaveChanges();
                }
            }
            catch (Exception e)
            {
                IntoFile("MachineID: " + machineId + ". InsetingDataIntoModeTable ::  " + e.ToString());
            }

            return await Task<int>.FromResult(Ret);
        }


        #region To update LossofEntry when  mode changes from IDLE to Production
        public async Task<int> UpdateModeRow(DateTime dt, int lastmachineid, int machineId, string PresentColor, string correctedDate)
        {
            int Ret = 1;
            tbllivemodedb tmprevious = new tbllivemodedb();
            tbllivelossofentry livelossofentryrow = new tbllivelossofentry();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                tmprevious = db.tbllivemodedbs.Find(lastmachineid);
                livelossofentryrow = db.tbllivelossofentries.Where(m => m.MachineID == machineId && m.CorrectedDate == correctedDate).OrderByDescending(m => m.LossID).FirstOrDefault();
            }

            DateTime ModeStartTime = Convert.ToDateTime(tmprevious.StartTime);
            tmprevious.EndTime = dt;
            tmprevious.IsCompleted = 1;
            double duationinsec = dt.Subtract(Convert.ToDateTime(tmprevious.StartTime)).TotalSeconds;
            tmprevious.DurationInSec = Convert.ToInt32(duationinsec);
            try
            {
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    db5.Entry(tmprevious).State = System.Data.Entity.EntityState.Modified;
                    db5.SaveChanges();
                }

                // While changing mode from  IDLE to Production
                if ((tmprevious.Mode == "IDLE") && (PresentColor == "green" || PresentColor == "blue"))
                {
                    if (livelossofentryrow.StartDateTime >= ModeStartTime)
                    {
                        livelossofentryrow.EndDateTime = dt;
                        livelossofentryrow.DoneWithRow = 1;
                        livelossofentryrow.IsUpdate = 1;
                        livelossofentryrow.IsScreen = 0;
                        livelossofentryrow.IsStart = 0;
                        livelossofentryrow.ForRefresh = 0;
                        using (i_facility_talEntities db1 = new i_facility_talEntities())
                        {
                            db1.Entry(livelossofentryrow).State = System.Data.Entity.EntityState.Modified;
                            db1.SaveChanges();
                        }
                    }
                }

            }
            catch (Exception e)
            {
                IntoFile("MachineID: " + machineId + ". InsetingDataIntoModeTable ::  " + e.ToString());
            }

            return await Task<int>.FromResult(Ret);
        }
        #endregion

        #region To update LossofEntry when  mode changes from IDLE to Production
        public async Task<int> UpdateModeRowWithStartTime(DateTime dt, int lastmodeID, int machineId, DateTime CompletedModeET, string PresentColor, string correctedDate)
        {
            int Ret = 1;
            tbllivemodedb tmlastmode = new tbllivemodedb();
            tbllivelossofentry livelossofentryrow = new tbllivelossofentry();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                tmlastmode = db.tbllivemodedbs.Find(lastmodeID);
                livelossofentryrow = db.tbllivelossofentries.Where(m => m.MachineID == machineId && m.CorrectedDate == correctedDate).OrderByDescending(m => m.LossID).FirstOrDefault();
            }
            //CompletedModeET = DateTime.Now;
            tmlastmode.StartTime = CompletedModeET;
            tmlastmode.EndTime = dt;
            tmlastmode.IsCompleted = 1;

            DateTime ModeStartTime = Convert.ToDateTime(tmlastmode.StartTime);
            // dt = DateTime.Now;
            IntoFile("ENDTIME:" + CompletedModeET);
            double duationinsec = dt.Subtract(Convert.ToDateTime(CompletedModeET)).TotalSeconds;
            IntoFile("Duratin:" + duationinsec);
            tmlastmode.DurationInSec = Convert.ToInt32(duationinsec);
            try
            {
                using (i_facility_talEntities db = new i_facility_talEntities())
                {
                    db.Entry(tmlastmode).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }


                // While changing mode from  IDLE to Production
                if ((tmlastmode.Mode == "IDLE") && (PresentColor == "green" || PresentColor == "blue"))
                {
                    if (livelossofentryrow.StartDateTime <= ModeStartTime)
                    {
                        livelossofentryrow.EndDateTime = dt;
                        livelossofentryrow.DoneWithRow = 1;
                        livelossofentryrow.IsUpdate = 1;
                        livelossofentryrow.IsScreen = 0;
                        livelossofentryrow.IsStart = 0;
                        livelossofentryrow.ForRefresh = 0;
                        using (i_facility_talEntities db1 = new i_facility_talEntities())
                        {
                            db1.Entry(livelossofentryrow).State = System.Data.Entity.EntityState.Modified;
                            db1.SaveChanges();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                IntoFile("AddModeRow Method For MachineID:" + machineId + "Error is" + e.ToString());
            }

            return await Task<int>.FromResult(Ret);
        }

        #endregion

        public async Task<int> UpdateModeRowWithStartTime(DateTime dt, int lastmodeID, int machineId, DateTime CompletedModeET)
        {
            int Ret = 1;
            tbllivemodedb tmlastmode = new tbllivemodedb();
            using (i_facility_talEntities db = new i_facility_talEntities())
            {
                tmlastmode = db.tbllivemodedbs.Find(lastmodeID);
            }
            //CompletedModeET = DateTime.Now;
            tmlastmode.StartTime = CompletedModeET;
            tmlastmode.EndTime = dt;
            tmlastmode.IsCompleted = 1;

            // dt = DateTime.Now;
            IntoFile("ENDTIME:" + CompletedModeET);
            double duationinsec = dt.Subtract(Convert.ToDateTime(CompletedModeET)).TotalSeconds;
            IntoFile("Duratin:" + duationinsec);
            tmlastmode.DurationInSec = Convert.ToInt32(duationinsec);
            try
            {
                using (i_facility_talEntities db = new i_facility_talEntities())
                {
                    db.Entry(tmlastmode).State = System.Data.Entity.EntityState.Modified;
                    db.SaveChanges();
                }
            }
            catch (Exception e)
            {
                IntoFile("AddModeRow Method For MachineID:" + machineId + "Error is" + e.ToString());
            }

            return await Task<int>.FromResult(Ret);
        }

        private void MyTimer_ForDeletingFromPCBDAQINTable(object sender, EventArgs e)
        {
            Method_ForDeleting();
        }

        //public void Method_ForDeleting()
        //{
        //    //NEW Delete Logic 2016-11-03

        //    DateTime DeleteDataBeforeThisDateTime = DateTime.Now.AddMinutes(-3);
        //    ConnectionString mc = null;
        //    try
        //    {
        //        mc = new ConnectionString();
        //        mc.open();
        //        //SqlCommand cmdDeleteRows = new SqlCommand("delete from [i_facility_tsal].[dbo].pcbdaqin_tblNew where DAQINID NOT IN ( " +
        //        //                                                " Select DAQINID from ( " +
        //        //                                                " Select DAQINID from ( " +
        //        //                                                " select DAQINID ,v.MachineID, d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn " +
        //        //                                                " from pcbdaqin_tblNew as d , VW_Join_DetailsANDParam as v " +
        //        //                                                " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc ) as T " +
        //        //                                                " group by T.PCBIPAddress,T.ParamPIN ) as k) and CreatedOn < '" + DeleteDataBeforeThisDateTime.ToString("yyyy-MM-dd HH:mm:00") + "' ;", mc.msqlConnection);

        //        //SqlCommand cmdDeleteRows = new SqlCommand("delete from [i_facility_tsal].[dbo].pcbdaqin_tblNew where DAQINID NOT IN(" +
        //        //                  "select DAQINID from pcbdaqin_tblNew as d , VW_Join_DetailsANDParam as v where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN)", mc.msqlConnection);
        //        //cmdDeleteRows.ExecuteNonQuery();
        //        mc.close();
        //        DataTable dataHolderj = new DataTable();
        //        string sql1 = "Select d.DAQINID ,v.MachineID , d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn from [i_facility_tsal].[dbo].pcbdaqin_tblNew as d , vw_join_detailsandparamNew as v "
        //                      + " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc";

        //        using (SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection))
        //        {
        //            da1.Fill(dataHolderj);
        //        }



        //        DataView dt1 = new DataView(dataHolderj);
        //        dt1.Sort = "PCBIPAddress ASC,ParamPIN ASC";
        //        DataTable dtnew = dt1.ToTable();

        //        var qryLatestInterview = from rows in dtnew.AsEnumerable()
        //                                 group rows by new { PositionID = rows["PCBIPAddress"], CandidateID = rows["ParamPIN"] } into grp
        //                                 select grp.First();
        //        dtnew = qryLatestInterview.CopyToDataTable();
        //        int count = dtnew.Rows.Count;
        //        var ids = (from r in dtnew.AsEnumerable()
        //                   where r.Field<DateTime>("CreatedOn") < DeleteDataBeforeThisDateTime
        //                   select r.Field<int>("DAQINID")).ToList<int>();


        //    }
        //    catch (Exception ex)
        //    {
        //        IntoFile(" Delete PrevData Error Inside: " + ex.ToString());
        //    }
        //    finally
        //    {
        //        mc.close();
        //    }

        //}

        public void Method_ForDeleting()
        {
            //NEW Delete Logic 2016-11-03

            //DateTime DeleteDataBeforeThisDateTime = DateTime.Now.AddMinutes(-3);
            string DeleteDataBeforeThis = DateTime.Now.AddMinutes(-5).ToString("yyyy-MM-dd HH:mm:00");
            DateTime DeleteDataBeforeThisDateTime = Convert.ToDateTime(DeleteDataBeforeThis);
            ConnectionString mc = null;
            try
            {
                mc = new ConnectionString();
                mc.open();
                //SqlCommand cmdDeleteRows = new SqlCommand("delete from [i_facility_tsal].[dbo].pcbdaqin_tbl where DAQINID NOT IN ( " +
                //                                                " Select DAQINID from ( " +
                //                                                " Select DAQINID from ( " +
                //                                                " select DAQINID ,v.MachineID, d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn " +
                //                                                " from pcbdaqin_tbl as d , VW_Join_DetailsANDParam as v " +
                //                                                " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc ) as T " +
                //                                                " group by T.PCBIPAddress,T.ParamPIN ) as k) and CreatedOn < '" + DeleteDataBeforeThisDateTime.ToString("yyyy-MM-dd HH:mm:00") + "' ;", mc.msqlConnection);

                //SqlCommand cmdDeleteRows = new SqlCommand("delete from [i_facility_tsal].[dbo].pcbdaqin_tbl where DAQINID NOT IN(" +
                //                  "select DAQINID from pcbdaqin_tbl as d , VW_Join_DetailsANDParam as v where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN)", mc.msqlConnection);
                //cmdDeleteRows.ExecuteNonQuery();
                DataTable dataHolderj = new DataTable();
                string sql1 = "Select d.DAQINID ,v.MachineID , d.PCBIPAddress , d.ParamPIN , ParamValue , CreatedOn from [i_facility_tsal].[dbo].pcbdaqin_tblNew as d , vw_join_detailsandparamNew as v "
                              + " where v.PCBIPAddress = d.PCBIPAddress and v.PinNumber = d.ParamPIN  order by CreatedOn desc";

                using (SqlDataAdapter da1 = new SqlDataAdapter(sql1, mc.msqlConnection))
                {
                    da1.Fill(dataHolderj);
                }



                DataView dt1 = new DataView(dataHolderj);
                dt1.Sort = "PCBIPAddress ASC,ParamPIN ASC";
                DataTable dtnew = dt1.ToTable();

                IEnumerable<DataRow> qryLatestInterview = from rows in dtnew.AsEnumerable()
                                                          group rows by new { PositionID = rows["PCBIPAddress"], CandidateID = rows["ParamPIN"] } into grp
                                                          select grp.First();
                dtnew = qryLatestInterview.CopyToDataTable();
                int count = dtnew.Rows.Count;
                List<int> ids = (from r in dtnew.AsEnumerable()
                                 where r.Field<DateTime>("CreatedOn") < DeleteDataBeforeThisDateTime
                                 select r.Field<int>("DAQINID")).ToList<int>();
                string dt = string.Join(",", ids);

                SqlCommand cmdDeleteRows = new SqlCommand("delete from [i_facility_tsal].[dbo].pcbdaqin_tblNew where DAQINID NOT IN ( " + dt + ")", mc.msqlConnection);
                cmdDeleteRows.ExecuteNonQuery();

                //using (i_facility_talEntities db = new i_facility_talEntities())
                //{
                //    var pcbipList = (from daq in db.pcbdaqin_tblNew
                //                     select daq.PCBIPAddress).Distinct().ToList();

                //    var pcbpinList = (from daq in db.pcbdaqin_tblNew
                //                      select daq.ParamPIN).Distinct().ToList();

                //    if (pcbpinList.Count > 0)
                //    {
                //        foreach (var data in pcbipList)
                //        {
                //            foreach (var item in pcbpinList)
                //            {
                //                var ls = db.pcbdaqin_tblNew.Where(m => m.PCBIPAddress == data && m.ParamPIN == item).OrderByDescending(m => m.DAQINID).Skip(1).ToList();
                //                db.pcbdaqin_tblNew.RemoveRange(ls);
                //                db.SaveChanges();
                //            }
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                IntoFile(" Delete PrevData Error Inside: " + ex.ToString());
            }
            finally
            {
                mc.close();
            }

        }

        public void IntoFile(string Msg)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Application.StartupPath + @"\Tata_UAT_ProductionStatusLogFile_"+DateTime.Now.ToString("yyyy-MM-dd")+".txt";
            using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
            {
                writer.WriteLine(System.DateTime.Now + ":  " + Msg);
            }
        }

        public void IntoFile1(string Msg)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Application.StartupPath + @"\PINGLogFile" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
            {
                writer.WriteLine(System.DateTime.Now + ":  " + Msg);
            }
        }

        public string GetCorrecteddate(string CorrectedDate)
        {
            string result = "Nochnage";
            DateTime DayTimingDet = DateTime.Now;
            if (CorrectedDate == Convert.ToString(DayTimingDet.Date))
            {
                if (DayTimingDet.Hour < 6)
                {

                }
            }

            //try
            //{
            //    using (i_facility_talEntities db = new i_facility_talEntities())
            //    {
            //        DAyTimingDet = Convert.ToDateTime(CorrectedDate + " " + Convert.ToString(db.tbldaytimings.Where(m => m.IsDeleted == 0).Select(m => m.EndTime).FirstOrDefault()));
            //    }
            //}
            //catch (Exception e)
            //{

            //}
            //if (DAyTimingDet != null)
            //{
            //    DateTime TodayDate = DateTime.Now;
            //    if (TodayDate < DAyTimingDet)
            //    {
            //        result = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            //    }
            //    else
            //    {
            //        result = DateTime.Now.ToString("yyyy-MM-dd");
            //    }
            //}
            return result;
        }

        private void Form1_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            // if not restarting the uncomment the line of code which is now commented
            //System.Diagnostics.Process.Start(Application.StartupPath);
            Application.Restart();
        }

        public async Task<bool> DAQ( List<tblmachinedetail> machineList, DataTable dataHolderj,string CorrectedDate,string Shift)
        {
            bool isdone = false;
            foreach (tblmachinedetail machine in machineList)
            {
                int machineId = machine.MachineID;
                int isfourthGen = machine.IsDLVersion;
                //dont push if last mode is Breakdown
                int IsBreakdown = 0;
                //IntoFile("MachineID:" + machineId + " MachineIPAddress :" + machine.IPAddress);
                // var lastMode = new tbllivemodedb();
                using (i_facility_talEntities db5 = new i_facility_talEntities())
                {
                    //lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).OrderByDescending(m => m.InsertedOn).Take(1).SingleOrDefault();
                    tbllivemodedb lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.ModeID).FirstOrDefault();
                    if (lastMode != null)
                    {
                        string mode = lastMode.Mode;
                        int modeStatus = lastMode.IsCompleted;
                        if (mode == "BREAKDOWN")
                        {
                            IsBreakdown = 1;
                        }
                        if (mode == "MNT")
                        {
                            IsBreakdown = 1;
                        }
                    }
                    else
                    {
                        IsBreakdown = 0;
                    }
                }


                if (IsBreakdown == 0)
                {
                    string pcbipaddress = "";
                    List<pcb_parameters> ParamList = new List<pcb_parameters>();
                    using (i_facility_talEntities db5 = new i_facility_talEntities())
                    {
                        //string pcbipaddress = db.pcb_details.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).Select(m => m.PCBIPAddress).FirstOrDefault();
                        pcbipaddress = db5.pcb_details.Where(m => m.IsDeleted == 0 && m.MachineID == machineId).Select(m => m.PCBIPAddress).FirstOrDefault();
                        ParamList = db5.pcb_parameters.Where(m => m.MachineID == machineId && m.IsDeleted == 0 && m.ParameterType == "PON").ToList();

                        //IntoFile("MachineID:" + machineId + " PCBIPAddress :" + pcbipaddress);
                    }
                    int count = 0;

                    TryPing:
                    bool pingable = false;
                    Ping pinger = new Ping();
                    try
                    {

                        PingReply reply = pinger.Send(pcbipaddress, 2000); //2000
                        pingable = reply.Status == IPStatus.Success;
                        //IntoFile1("before fourth gen Pinging Status of " + pcbipaddress + ": " + pingable);

                        // Removed the commented
                        if (isfourthGen == 1)
                        {
                            IntoFile1("Pinging Status of " + pcbipaddress + ": " + pingable);
                            pingable = true;
                        }

                        IntoFile("Pinging Status of " + pcbipaddress + ": " + pingable);

                        if (!pingable)
                        {
                            IntoFile1("Pinging Status of " + pcbipaddress + ": " + pingable);
                        }
                    }
                    catch (PingException e)
                    {
                        // Discard PingExceptions and return false;
                        IntoFile(" PING Exception" + "MachineID: " + machineId + " " + e.ToString());
                    }
                    catch (Exception e)
                    {
                        IntoFile(" During Ping " + "MachineID: " + machineId + " " + e.ToString());
                    }


                    if (pingable)
                    {
                        #region OLD NO-PING NO-BLUE
                        handle_no_ping NoPingData = new handle_no_ping();
                        using (i_facility_talEntities db5 = new i_facility_talEntities())
                        {
                            NoPingData = db5.handle_no_ping.Where(m => m.MachineID == machineId).FirstOrDefault();
                        }
                        if (NoPingData != null)
                        {
                            int id = NoPingData.NoPingID;
                            handle_no_ping hnp = db.handle_no_ping.Find(id);
                            db.handle_no_ping.Remove(hnp);
                            db.SaveChanges();
                        }
                        #endregion
                        bool ispacketloss = false;
                        IOTGatwayPacketsData IOTgatewaydata = db.IOTGatwayPacketsDatas.Where(m => m.IPAddres == pcbipaddress).OrderByDescending(m => m.GatewayMsgID).FirstOrDefault();

                        if (IOTgatewaydata != null)
                        {
                            DateTime DBdatetime = Convert.ToDateTime(IOTgatewaydata.CreatedOn);
                            DateTime CurrentTime = DateTime.Now;
                            double diff = CurrentTime.Subtract(DBdatetime).TotalMinutes;
                            if (diff >= 3)
                            {
                                ispacketloss = true;
                            }
                        }

                        if (!ispacketloss)
                        {
                            //Delete the Machine No Ping details from the DB.
                            PINGHANDLER_tbl dldata1 = db.PINGHANDLER_tbl.Where(m => m.MachineID == machineId && m.IPADDRESS == pcbipaddress).OrderByDescending(m => m.PID).FirstOrDefault();
                            if (dldata1 != null)
                            {
                                db.PINGHANDLER_tbl.Remove(dldata1);
                            }
                            List<DataRow> dr = dataHolderj.Select("MachineID = '" + @machineId + "'").ToList();
                            if (dr != null)
                            {
                                //Set Default values to All Pins
                                int pin17 = 0, pin18 = 0, pin19 = 0, pin20 = 0, pin22 = 0, pin23 = 0, pin24 = 0, pin25 = 0, pin26 = 0;
                                List<KeyValuePair<int, int>> Factorlist = new List<KeyValuePair<int, int>>();
                                //list.Add(new KeyValuePair<string, int>("Rabbit", 4));
                                #region Try Catch 1
                                try
                                {
                                    foreach (DataRow rowj in dr)
                                    {
                                        //var jack = rowj;
                                        int pinNo = Convert.ToInt32(rowj[3]);
                                        int pinVal = Convert.ToInt32(rowj[4]);
                                        IEnumerable<pcb_parameters> paramData = (from row in ParamList where row.PinNumber == pinNo select row);
                                        if (paramData != null)
                                        {
                                            Factorlist.Add(new KeyValuePair<int, int>(pinNo, pinVal));
                                        }

                                        //Now Assign Value taken from pcbdaqin_tblNew to PIN's
                                        switch (pinNo)
                                        {
                                            case 17:
                                                {
                                                    pin17 = pinVal;
                                                    break;
                                                }
                                            case 18:
                                                {
                                                    pin18 = pinVal;
                                                    break;
                                                }
                                            case 19:
                                                {
                                                    pin19 = pinVal;
                                                    break;
                                                }
                                            case 20:
                                                {
                                                    pin20 = pinVal;
                                                    break;
                                                }
                                            case 22:
                                                {
                                                    pin22 = pinVal;
                                                    break;
                                                }
                                            case 23:
                                                {
                                                    pin23 = pinVal;
                                                    break;
                                                }
                                            case 24:
                                                {
                                                    pin24 = pinVal;
                                                    break;
                                                }
                                            case 25:
                                                {
                                                    pin25 = pinVal;
                                                    break;
                                                }
                                            case 26:
                                                {
                                                    pin26 = pinVal;
                                                    break;
                                                }
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    IntoFile(" Gen KeyValuePair " + "MachineID: " + machineId + " " + e.ToString());
                                    //MessageBox.Show("Gen KeyValuePair " + e);
                                }

                                #endregion //try catch 1

                                #region Try Catch 2
                                try
                                {
                                    //int PONHighValue = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.HighValue).SingleOrDefault());
                                    //int PONPinNo = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.PinNumber).SingleOrDefault());
                                    int PONHighValue = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.HighValue).FirstOrDefault());
                                    int PONPinNo = Convert.ToInt32((from row in ParamList where row.ColorCode == "blue" select row.PinNumber).FirstOrDefault());

                                    List<int> ponvalue = (from kvp in Factorlist where kvp.Key == PONPinNo select kvp.Value).ToList();

                                    if (ponvalue.Count > 0)
                                    {
                                        int PONValue = ponvalue[0];

                                        if (PONHighValue == PONValue)
                                        {
                                            #region //New Logic to Decide the Color( Row ) into  tbllivemodedb Table.
                                            string ColorIntoTblMode = "yellow";
                                            pcbdps_master ColorData = new pcbdps_master();
                                            using (i_facility_talEntities db5 = new i_facility_talEntities())
                                            {
                                                ColorData = db.pcbdps_master.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.Pin17 == pin17 && m.Pin18 == pin18 && m.Pin19 == pin19 && m.Pin20 == pin20 && m.Pin22 == pin22 && m.Pin23 == pin23 && m.Pin24 == pin24 && m.Pin25 == pin25 && m.Pin26 == pin26).FirstOrDefault();
                                            }
                                            if (ColorData != null)
                                            {
                                                ColorIntoTblMode = ColorData.ColorValue;
                                            }
                                            if (ColorIntoTblMode == "yellow")
                                            {
                                                await InsertingDataIntoModeTable("IDLE", machineId, ColorIntoTblMode, CorrectedDate, Shift);
                                            }
                                            else if (ColorIntoTblMode == "green")
                                            {
                                                await InsertingDataIntoModeTable("PowerOn", machineId, ColorIntoTblMode, CorrectedDate, Shift);
                                            }
                                            #endregion
                                        }
                                        else
                                        {
                                            #region Check in tbllivemodedb and push. TODO

                                            tbllivemodedb lastMode = new tbllivemodedb();

                                            using (i_facility_talEntities db5 = new i_facility_talEntities())
                                            {
                                                lastMode = db5.tbllivemodedbs.Where(m => m.IsDeleted == 0 && m.MachineID == machineId && m.IsCompleted == 0).OrderByDescending(m => m.InsertedOn).FirstOrDefault();
                                            }
                                            if (lastMode != null)
                                            {
                                                string mode = lastMode.Mode;
                                                int modeStatus = lastMode.IsCompleted;
                                                if (mode == "BREAKDOWN")
                                                {
                                                    //await InsertingDataIntoModeTable("BREAKDOWN", machineId, "red", CorrectedDate, Shift);
                                                }
                                                else if (mode == "MNT")
                                                {
                                                    //await InsertingDataIntoModeTable("MNT", machineId, "red", CorrectedDate, Shift);
                                                }
                                                else
                                                {
                                                    await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                                                }
                                            }
                                            else
                                            {
                                                await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    IntoFile(" After KeyGen " + "MachineID: " + machineId + " " + e.ToString());
                                    //MessageBox.Show("After KeyGen: " + e);
                                }
                                #endregion  //end of try catch 2
                            }
                        }
                        else
                        {
                            await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                        }
                    }
                    else //if not pingable
                    {
                        count++;
                        ////Get Db Counter for No Pingable 4th Gen DL
                        //PINGHANDLER_tbl dldata = new PINGHANDLER_tbl();
                        //int DlCounter = 0;
                        //using (i_facility_talEntities db = new i_facility_talEntities())
                        //{
                        //    dldata = db.PINGHANDLER_tbl.Where(m => m.MachineID == machineId).OrderByDescending(m => m.PID).FirstOrDefault();
                        //}
                        //if (dldata != null)
                        //{
                        //    DlCounter = Convert.ToInt32(dldata.Pingcount);
                        //}


                        if (count < 2)//4
                        {
                            goto TryPing;
                        }
                        //else if (isfourthGen == 1 && DlCounter <= 6)
                        //{
                        //    DlCounter += 1;
                        //    //Insert or update into DB
                        //    if (dldata != null)
                        //    {
                        //        dldata.MachineID = machineId;
                        //        dldata.IPADDRESS = pcbipaddress;
                        //        dldata.CorrectedDate = CorrectedDate;
                        //        dldata.Pingcount = 1;
                        //        try
                        //        {
                        //            using (i_facility_talEntities db5 = new i_facility_talEntities())
                        //            {
                        //                db5.PINGHANDLER_tbl.Add(dldata);
                        //                db5.SaveChanges();
                        //            }
                        //        }
                        //        catch (Exception)
                        //        {
                        //            //IntoFile("AddModeRow Method For MachineID:" + machineId + "Error is" + e.ToString());
                        //        }
                        //    }
                        //    else
                        //    {
                        //        dldata.Pingcount += 1;
                        //        using (i_facility_talEntities db5 = new i_facility_talEntities())
                        //        {
                        //            db5.Entry(dldata).State = System.Data.Entity.EntityState.Modified;
                        //            db5.SaveChanges();
                        //        }
                        //    }

                        //    goto TryPing;    // Mode will delay's due to this
                        //}

                        else
                        {
                            //if (isfourthGen == 1)
                            //{
                            //    //pingable = true;
                            //}
                            //else
                            await InsertingDataIntoModeTable("PowerOff", machineId, "blue", CorrectedDate, Shift);
                        }
                    }
                }
            }
            return await Task<bool>.FromResult(isdone);
        }
    }
}
