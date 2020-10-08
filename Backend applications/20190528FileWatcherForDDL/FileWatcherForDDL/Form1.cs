using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;


namespace FileWatcherForDDL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            try
            {
                i_facility_talEntities db = new i_facility_talEntities();
                
                var nwDetails = db.tblnetworkdetailsforddls.Where(m => m.IsDeleted == 0).FirstOrDefault();
               string path = nwDetails.Path;
               // string path = @"D:\PCS-DAS-Sync\DDL";
                string username = nwDetails.UserName;
                string password = nwDetails.Password;
                string domainname = nwDetails.DomainName;

               var nwDetails1 = db.tblnetworkdetailsforddls.Where(m => m.IsDeleted == 0 && m.NPFDDLID == 4).FirstOrDefault(); // INCremental DDL
                string path1 = nwDetails1.Path;
                //string path1 = @"D:\PCS-DAS-Sync\INCDDL";
                string username1 = nwDetails1.UserName;
                string password1 = nwDetails1.Password;
                string domainname1 = nwDetails1.DomainName;
                try
                {
                    //InitializeComponent();
                    IntoFile("ApplicationStart");
                    InitializeComponent(path, username, password, domainname, path1, username1, password1, domainname1);
                }
                catch (Exception e)
                {
                    //MessageBox.Show("Path Error: " + e);
                    IntoFile(e.ToString());
                }

                //string destinationPathWithFileName = @"C:\Users\TECH-1\Desktop\Tsal Unit work centres_10.04.2017.xlsx";
                //HandleDDLFile(destinationPathWithFileName);
            }
            catch (Exception exc)
            {
                //MessageBox.Show("Path Error: " + e);
                IntoFile(exc.ToString());
            }
        }

       private void fileSystemWatcher1_Created(object sender, FileSystemEventArgs e)
        {
            IntoFile("File watcher method");
            try
            {
                string rootFolderPath = e.FullPath;
                string filename = Path.GetFileName(rootFolderPath);
                IntoFile("File watcher method filename:" + filename);
                string fileExt = Path.GetExtension(rootFolderPath);
                if ((!filename.Contains("~$")) && (fileExt == ".xlsx"))
                {
                    string destinationPathWithFileName = @"\\10.80.20.16\DDL-DAS Report\Uploaded DDL\" + filename;

                    //string destinationPathWithFileName = @"D:\PCS-DAS-Sync\Uploaded DDL\" + filename;
                    IntoFile("File watcher method destination path:" + destinationPathWithFileName);

                    

                    Thread.Sleep(1 *60* 1000);
                    File.Copy(rootFolderPath, destinationPathWithFileName, true);

                    //Delete the Original File.
                    try
                    {
                        if (File.Exists(rootFolderPath))
                        {
                            File.Delete(rootFolderPath);
                        }
                    }
                    catch (Exception eD)
                    {
                        IntoFile("Deleting Original File: " + eD);
                    }

                    try
                    {
                        IntoFile("File watcher method before handleddlfile:");
                        HandleDDLFile(destinationPathWithFileName);
                    }
                    catch (Exception ex)
                    {
                        IntoFile(ex.ToString());
                    }
                    finally
                    {
                    }

                    //Delete Duplicates by calling StoredProcedure
                    try
                    {
                        MsqlConnection con = new MsqlConnection();
                        //string conString = "server = 'localhost' ;userid = 'root' ;Password = 'srks4$' ;database = 'mazakdaq';port = 3306 ;persist security info=False";
                        //string conString = "Data Source = DESKTOP-72HGDFG\\SQLDEV17013; User ID = sa;Password = srks4$;Initial Catalog = i_facility_tal;Persist Security Info=True";
                        using (SqlConnection databaseConnection = new SqlConnection(con.msqlConnection.ConnectionString))
                        {
                            SqlCommand cmd = new SqlCommand("DeleteDupWorkOrdersFromDDLList", databaseConnection);
                            databaseConnection.Open();
                            cmd.CommandType = CommandType.StoredProcedure;
                            var a = cmd.ExecuteNonQuery();
                            //MessageBox.Show(a.ToString());
                            databaseConnection.Close();
                        }
                    }
                    catch (Exception DelSP)
                    {
                        IntoFile("DeleteDuplicate SP Error: " + DelSP);
                    }
                }
                else
                {
                    // file Error
                    IntoFile("File Format Error ");
                }
            }
            catch (Exception exc)
            {
                IntoFile("IN FileWatcher Section: " + exc);
            }

            //Commented by vignesh no need to close and start the application here written new code has form closed event

            //System.Diagnostics.Process.Start(Application.ExecutablePath); // to start new instance of application
            //this.Close(); //to turn off current app
        }

       private void fileSystemWatcher2_Created(object sender, FileSystemEventArgs e)
        {
            try
            {
                string rootFolderPath = e.FullPath;
                string filename = Path.GetFileName(rootFolderPath);
                string fileExt = Path.GetExtension(rootFolderPath);
                if ((!filename.Contains("~$")) && (fileExt == ".xlsx"))
                {
                    string destinationPathWithFileName = @"\\10.80.20.16\DDL-DAS Report\Uploaded Incremental DDL\" + filename;
                    //string destinationPathWithFileName = @"D:\PCS-DAS-Sync\Uploaded Incremental DDL\" + filename;

                    Thread.Sleep(1 * 60 * 1000);
                    File.Copy(rootFolderPath, destinationPathWithFileName, true);

                    //Delete the Original File.
                    try
                    {
                        if (File.Exists(rootFolderPath))
                        {
                            File.Delete(rootFolderPath);
                        }
                    }
                    catch (Exception eD)
                    {
                        IntoFile("Deleting Original File: " + eD);
                    }

                    try
                    {
                        HandleDDLFileincemental(destinationPathWithFileName);
                    }
                    catch (Exception ex)
                    {
                        IntoFile(ex.ToString());
                    }
                    finally
                    {
                    }
                    #region Added by AShok
                    //Delete Duplicates by calling StoredProcedure
                    try
                    {
                        MsqlConnection con = new MsqlConnection();
                        //string conString = "server = 'localhost' ;userid = 'root' ;Password = 'srks4$' ;database = 'mazakdaq';port = 3306 ;persist security info=False";
                        //string conString = "Data Source = SRKS - TECH3 - PC\\SQLEXPRESS; User ID = sa;Password = srks4$;Initial Catalog = i_facility_tsal;Persist Security Info=True";
                        using (SqlConnection databaseConnection = new SqlConnection(con.msqlConnection.ConnectionString))
                        {
                            SqlCommand cmd = new SqlCommand("DeleteDupWorkOrdersFromDDLList", databaseConnection);
                            databaseConnection.Open();
                            cmd.CommandType = CommandType.StoredProcedure;
                            var a = cmd.ExecuteNonQuery();
                            //MessageBox.Show(a.ToString());
                            databaseConnection.Close();
                        }
                    }
                    catch (Exception DelSP)
                    {
                        IntoFile("DeleteDuplicate SP Error: " + DelSP);
                    }
                    #endregion
                }
                else
                {
                    // file Error
                    IntoFile("File Format Error ");
                }
            }
            catch (Exception exc)
            {
                IntoFile("IN FileWatcher Section: " + exc);
            }


            //Commented by vignesh no need to close and start the application here written new code has form closed event

            //System.Diagnostics.Process.Start(Application.ExecutablePath); // to start new instance of application
            //this.Close(); //to turn off current app
        }

        public void HandleDDLFile(string DDLPath)
        {
            try
            {
                IntoFile("HandleDDLFile:" + DDLPath);
               // CheckDDLStatus(0);// insert the row from ddl status table
                DataSet ds = new DataSet();
                FileInfo f = new FileInfo(DDLPath);

                #region SHIFT and CorrectedDate

                string Shift = null;
                DataTable dtshift = new DataTable();
                String queryshift = "SELECT ShiftName,StartTime,EndTime FROM shift_master WHERE IsDeleted = 0";
                MsqlConnection mcp = new MsqlConnection();
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

                if (f.Length > 0)
                {
                    //string fileExtension = System.IO.Path.GetExtension(Request.Files["file"].FileName);
                    string fileExtension = Path.GetExtension(DDLPath);
                    string filename = Path.GetFileName(DDLPath);
                    string fileLocation = AppDomain.CurrentDomain.BaseDirectory + "\\DDL\\" + filename;
                    IntoFile("HandleDDLFile fileLocation:" + fileLocation);
                    //string folderPath = @"C:\UnitWorks\Testing FileWatcher\Debug";
                    //fileLocation = @"C:\UnitWorks\Testing FileWatcher\Debug" + "\\DDL\\" + filename;

                    #region  //Deleting Excel file
                    //DirectoryInfo di = new DirectoryInfo(folderPath);
                    //FileInfo[] files = di.GetFiles("*.xlsx").Where(p => p.Extension == ".xlsx").ToArray();
                    //foreach (FileInfo file1 in files)
                    //    try
                    //    {
                    //        file1.Attributes = FileAttributes.Normal;
                    //        System.IO.File.Delete(file1.FullName);
                    //    }
                    //    catch { }
                    #endregion

                    try
                    {
                        if (File.Exists(fileLocation))
                        {
                            File.Delete(fileLocation);
                        }
                        Thread.Sleep(1 * 60 * 1000);
                        File.Copy(DDLPath, fileLocation);
                    }
                    catch (Exception fe)
                    {
                        IntoFile(fe.ToString());
                    }

                    DataTable dt = new DataTable();
                    try
                    {
                        Thread.Sleep(1 * 60 * 1000);
                        dt = GetDataTableFromExcel(fileLocation);
                    }
                    catch (Exception exReadExcel)
                    {
                        IntoFile(exReadExcel.ToString());
                    }
                    
                    try
                    {
                        MsqlConnection mcDeleteRows = new MsqlConnection();
                        mcDeleteRows.open();

                        //String DeleteQuery = "set sql_safe_updates = 0;" +
                        //    "DELETE from i_facility_tal.dbo.tblddl where IsCompleted = 0 or CorrectedDate <= '" + System.DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd") + "' ;" +
                        //    "set sql_safe_updates = 1;";

                        String DeleteQuery = "DELETE from " + MsqlConnection.DBSchemaName + ".tblddl where IsCompleted = 0 or CorrectedDate <= '" + System.DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd") + "' ;";

                        //MySqlCommand cmdDeleteRows = new MySqlCommand("TRUNCATE TABLE mazakdaq.tblddl;", mcDeleteRows.msqlConnection);
                        SqlCommand cmdDeleteRows = new SqlCommand(DeleteQuery, mcDeleteRows.msqlConnection);
                        cmdDeleteRows.ExecuteNonQuery();
                        mcDeleteRows.close();
                    }
                    catch (Exception e)
                    {
                        //MessageBox.Show("Delete Error:");
                        IntoFile("Delete Error " + e.ToString());
                    }
                    
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            string a = dt.Rows[i][0].ToString();

                            try
                            {
                                string flagrush = null;
                                flagrush = Convert.ToString(dt.Rows[i][14]);
                                if (flagrush.Length > 0)
                                {
                                    flagrush = flagrush.Trim();
                                }
                                if (flagrush == "7" || flagrush == "9")
                                {
                                    continue;
                                }
                                else
                                {

                                    //MessageBox.Show(" No Error in Opening " + a);

                                    string dat = DateTime.Now.ToString();
                                    dat = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


                                    // //MessageBox.Show("0 in excel. " + dt.Rows[i][0].ToString());
                                    string WorkCenter = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][0].ToString()))
                                    {
                                        try
                                        {
                                            WorkCenter = @dt.Rows[i][0].ToString();
                                            if (WorkCenter.Length > 0)
                                            {
                                                WorkCenter = WorkCenter.Trim();
                                            }
                                            //WorkCenter = RemoveSpecialCharacters(WorkCenter);
                                            WorkCenter = WorkCenter.Replace("'", "");
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    string WorkOrder = null;
                                    WorkOrder = Convert.ToString(dt.Rows[i][1]).Trim();

                                    string OpNo = null;
                                    OpNo = Convert.ToString(dt.Rows[i][2]).Trim();

                                    string OperationDesc = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][3])))
                                    {
                                        try
                                        {
                                            OperationDesc = @dt.Rows[i][3].ToString();
                                            //MessageBox.Show("Before" + OperationDesc);
                                            //OperationDesc = RemoveSpecialCharacters(OperationDesc); Replace("'", "");
                                            OperationDesc = OperationDesc.Replace("'", "");
                                            OperationDesc = OperationDesc.Replace("\"", "");
                                            if (OperationDesc.Length > 0)
                                            {
                                                OperationDesc = OperationDesc.Trim();
                                            }
                                            //MessageBox.Show("After" + OperationDesc);
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }


                                    }
                                    string partNo = null;
                                    partNo = Convert.ToString(dt.Rows[i][4]).Trim();

                                    string PartName = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][5])))
                                    {
                                        try
                                        {
                                            PartName = @dt.Rows[i][5].ToString();
                                            //PartName = RemoveSpecialCharacters(PartName);
                                            PartName = PartName.Replace("'", "");
                                            PartName = PartName.Replace("\"", "");
                                            if (PartName.Length > 0)
                                            {
                                                PartName = PartName.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string type = null;
                                    type = Convert.ToString(dt.Rows[i][6]);
                                    int TargetQty = 0;
                                    string TargQty = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][7])))
                                    {
                                        // //MessageBox.Show("T : " + dt.Rows[i][8].ToString());
                                        try
                                        {
                                            TargQty = Convert.ToString(dt.Rows[i][7]);
                                            TargQty = TargQty.Replace(",", "");
                                            string[] inta = TargQty.Trim().Split('.');
                                            TargetQty = Convert.ToInt32(inta[0]);
                                        }
                                        catch (Exception e)
                                        {
                                            IntoFile("" + e);
                                            continue;
                                        }

                                    }
                                    else
                                    {
                                        IntoFile("TargetQty Empty for" + WorkOrder + " OpNo:" + OpNo);
                                        continue;
                                    }
                                    if (TargetQty == 0)
                                    {
                                        IntoFile("TargetQty 0 for" + WorkOrder + " OpNo:" + OpNo);
                                        continue;
                                    }

                                    string madDate = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][8])))
                                    {
                                        ////MessageBox.Show("Col 9. Row : " + i + " " + dt.Rows[i][9].ToString());
                                        try
                                        {
                                            string[] words = dt.Rows[i][8].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                // ////MessageBox.Show("INtermeditate : " + IntermediateMadDate);
                                                madDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }

                                    //MessageBox.Show(WorkCenter + " " + WorkOrder + " " + OpNo + " " + OperationDesc + " " + partNo + " " + PartName + " " + type + " " + TargQty + " " + madDate);

                                    string madDateInd = null;
                                    try
                                    {
                                        madDateInd = Convert.ToString(dt.Rows[i][9]);
                                        if (madDateInd.Length > 0)
                                        {
                                            madDateInd = madDateInd.Trim();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        //continue;
                                        IntoFile(e.ToString());
                                    }


                                    string preopnEndDate = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][10].ToString()))
                                    {
                                        try
                                        {
                                            //MessageBox.Show("Col 11. Row : " + i + " " + dt.Rows[i][10].ToString());
                                            string[] words = dt.Rows[i][10].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                preopnEndDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");

                                            }
                                            if (preopnEndDate.Length > 0)
                                            {
                                                preopnEndDate = preopnEndDate.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string DaysAgeing = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][11].ToString()))
                                    {
                                        try
                                        {
                                            //DaysAgeing = Convert.ToString(dt.Rows[i][11]) != null ? Convert.ToString(dt.Rows[i][11]).Trim() : Convert.ToString(dt.Rows[i][11]);
                                            DaysAgeing = Convert.ToString(dt.Rows[i][11]);
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                        if (DaysAgeing.Length > 0)
                                        {
                                            DaysAgeing = DaysAgeing.Trim();
                                        }

                                    }
                                    string Project = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][12].ToString()))
                                    {
                                        try
                                        {
                                            Project = dt.Rows[i][12].ToString();
                                            if (Project.Length > 0)
                                            {
                                                Project = Project.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string dueDate = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][13].ToString()))
                                    {
                                        try
                                        {
                                            ////MessageBox.Show("Col 14. Row : " + i + " " + dt.Rows[i][14].ToString());
                                            string[] words = dt.Rows[i][13].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                ////MessageBox.Show("INtermeditate : " + IntermediateMadDate);
                                                dueDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");
                                            }
                                            if (dueDate.Length > 0)
                                            {
                                                dueDate = dueDate.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    flagrush = null;
                                    try
                                    {
                                        flagrush = Convert.ToString(dt.Rows[i][14]);
                                        if (flagrush.Length > 0)
                                        {
                                            flagrush = flagrush.Trim();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        //continue;
                                        IntoFile(e.ToString());
                                    }

                                    string OpOnHold = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][15].ToString()))
                                    {
                                        try
                                        {
                                            OpOnHold = dt.Rows[i][15].ToString();
                                            if (OpOnHold.Length > 0)
                                            {
                                                OpOnHold = OpOnHold.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }
                                    string OpOnHoldReason = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][16].ToString()))
                                    {
                                        try
                                        {
                                            OpOnHoldReason = dt.Rows[i][16].ToString();
                                            if (OpOnHoldReason.Length > 0)
                                            {
                                                OpOnHoldReason = OpOnHoldReason.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    string SplitWO = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][17].ToString()))
                                    {
                                        try
                                        {
                                            SplitWO = dt.Rows[i][17].ToString();
                                            if (SplitWO.Length > 0)
                                            {
                                                SplitWO = SplitWO.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }
                                    if ((!String.IsNullOrEmpty(WorkOrder)) && (!String.IsNullOrEmpty(OpNo)) && (!String.IsNullOrEmpty(partNo)))
                                    {
                                        try
                                        {
                                            using (MsqlConnection mc1 = new MsqlConnection())
                                            {
                                                mc1.open();
                                                SqlCommand cmd2 = null;
                                                //MessageBox.Show("WO & OPNo & PartNo Values: " + WorkOrder + " " + OpNo + " " + partNo);
                                                //cmd2 = new MySqlCommand("INSERT INTO tblmachinedetails (InsertedBy,InsertedOn,IsDeleted,MachineType, MachineInvNo, IPAddress, ControllerType,MachineModel,MachineMake,ModelType,MachineDispName,IsParameters,ShopNo,IsPCB) VALUES( '" + dat + "'," + 0 + "," + 2 + ",'" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','" + ds.Tables[0].Rows[i][5].ToString() + "','" + ds.Tables[0].Rows[i][6].ToString() + "','" + ds.Tables[0].Rows[i][7].ToString() + "','" + ds.Tables[0].Rows[i][8].ToString() + "','" + ds.Tables[0].Rows[i][9].ToString() + "')", mc1.msqlConnection);
                                                cmd2 = new SqlCommand("INSERT INTO " + MsqlConnection.DBSchemaName + ".tblddl(WorkCenter,WorkOrder,OperationNo,OperationDesc,MaterialDesc,PartName,Type,TargetQty,MADDate,MADDateInd,PreOpnEndDate,DaysAgeing,Project,DueDate,FlagRushInd,OperationsOnHold,ReasonForHold,SplitWO,InsertedOn,CorrectedDate)VALUES('" + WorkCenter + "','" + WorkOrder + "','" + OpNo + "','" + OperationDesc + "','" + partNo + "','" + PartName + "','" + type + "','" + TargetQty + "','" + madDate + "','" + madDateInd + "','" + preopnEndDate + "','" + DaysAgeing + "','" + Project + "','" + dueDate + "','" + flagrush + "','" + OpOnHold + "','" + OpOnHoldReason + "','" + SplitWO + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + CorrectedDate + "');", mc1.msqlConnection);
                                                cmd2.ExecuteNonQuery();
                                                mc1.close();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }

                                }
                            }
                            catch (Exception e)
                            {
                                //MessageBox.Show("Insert Module: " + e);
                                IntoFile(e.ToString());
                            }
                            //if (mc1 != null)
                            //{
                            //    try
                            //    {
                            //        mc1.close();
                            //    }catch( Exception exce)
                            //    {
                            //        IntoFile("" + exce);
                            //    }
                            //}
                        }
                        catch (Exception exd)
                        {
                            //MessageBox.Show("Error in for Loop " + exd);
                            IntoFile(exd.ToString());
                        }
                    }

                }
               // CheckDDLStatus(1);// delete the row from ddl status table
            }
            catch (Exception exc)
            {
                IntoFile("HandleDDLFile Section: " + exc);
            }

        }

        #region Checking the Progress

        //public void CheckDDLStatus(int i)
        //{
            
        //    if (i == 1)
        //    {
        //        string statusQuery = "Truncate table [i_facility_tal].[dbo].[tblDDLStatus]";
        //        MsqlConnection con = new MsqlConnection();
        //        con.open();
        //        using (SqlConnection databaseConnection = new SqlConnection(con.msqlConnection.ConnectionString))
        //        {
                   
        //            SqlCommand cmd = new SqlCommand(statusQuery, databaseConnection);
        //            cmd.ExecuteNonQuery();

        //        }
        //        con.close();
        //    }
        //    else
        //    {
        //        using (MsqlConnection mc1 = new MsqlConnection())
        //        {
        //            mc1.open();
        //             //getting error here bject set t a null so need to verify th
        //             string queryInsert = "INSERT INTO [i_facility_tal].[dbo].[tblDDLStatus] ([StatusMessage],[Status],[CreatedOn],[CreatedBy],[IsDeleted])VALUES('DDL file Uploading', 1, '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "', 1, 0)";
        //            SqlCommand cmd2 = new SqlCommand(queryInsert,mc1.msqlConnection);
        //            cmd2.ExecuteNonQuery();
        //            mc1.close();
        //        }
                
        //    }
        //}

        #endregion


        public void HandleDDLFileincemental(string DDLPath)
        {
            try
            {
               // CheckDDLStatus(0);
                DataSet ds = new DataSet();
                FileInfo f = new FileInfo(DDLPath);

                #region SHIFT and CorrectedDate

                string Shift = null;
                DataTable dtshift = new DataTable();
                String queryshift = "SELECT ShiftName,StartTime,EndTime FROM shift_master WHERE IsDeleted = 0";
                MsqlConnection mcp = new MsqlConnection();
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

                if (f.Length > 0)
                {
                    //string fileExtension = System.IO.Path.GetExtension(Request.Files["file"].FileName);
                    string fileExtension = Path.GetExtension(DDLPath);
                    string filename = Path.GetFileName(DDLPath);
                    string fileLocation = AppDomain.CurrentDomain.BaseDirectory + "\\INCDDL\\" + filename;
                    //string folderPath = @"C:\UnitWorks\Testing FileWatcher\Debug";
                    //fileLocation = @"C:\UnitWorks\Testing FileWatcher\Debug" + "\\INCDDL\\" + filename;
                    
                    try
                    {
                        if (File.Exists(fileLocation))
                        {
                            File.Delete(fileLocation);
                        }
                        Thread.Sleep(1 * 60 * 1000);
                        File.Copy(DDLPath, fileLocation);
                    }
                    catch (Exception fe)
                    {
                        IntoFile(fe.ToString());
                    }

                    DataTable dt = new DataTable();
                    try
                    {
                        Thread.Sleep(1 * 60 * 1000);
                        dt = GetDataTableFromExcel(fileLocation);
                    }
                    catch (Exception exReadExcel)
                    {
                        IntoFile(exReadExcel.ToString());
                    }
                    
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            string a = dt.Rows[i][0].ToString();

                            try
                            {
                                string flagrush = null;
                                flagrush = Convert.ToString(dt.Rows[i][14]);
                                if (flagrush.Length > 0)
                                {
                                    flagrush = flagrush.Trim();
                                }
                                if (flagrush == "7" || flagrush == "9")
                                {
                                    continue;
                                }
                                else
                                {

                                    //MessageBox.Show(" No Error in Opening " + a);

                                    string dat = DateTime.Now.ToString();
                                    dat = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


                                    // //MessageBox.Show("0 in excel. " + dt.Rows[i][0].ToString());
                                    string WorkCenter = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][0].ToString()))
                                    {
                                        try
                                        {
                                            WorkCenter = @dt.Rows[i][0].ToString();
                                            if (WorkCenter.Length > 0)
                                            {
                                                WorkCenter = WorkCenter.Trim();
                                            }
                                            //WorkCenter = RemoveSpecialCharacters(WorkCenter);
                                            WorkCenter = WorkCenter.Replace("'", "");
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    string WorkOrder = null;
                                    WorkOrder = Convert.ToString(dt.Rows[i][1]).Trim();

                                    string OpNo = null;
                                    OpNo = Convert.ToString(dt.Rows[i][2]).Trim();

                                    string OperationDesc = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][3])))
                                    {
                                        try
                                        {
                                            OperationDesc = @dt.Rows[i][3].ToString();
                                            //MessageBox.Show("Before" + OperationDesc);
                                            //OperationDesc = RemoveSpecialCharacters(OperationDesc); Replace("'", "");
                                            OperationDesc = OperationDesc.Replace("'", "");
                                            OperationDesc = OperationDesc.Replace("\"", "");
                                            if (OperationDesc.Length > 0)
                                            {
                                                OperationDesc = OperationDesc.Trim();
                                            }
                                            //MessageBox.Show("After" + OperationDesc);
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }


                                    }
                                    string partNo = null;
                                    partNo = Convert.ToString(dt.Rows[i][4]).Trim();

                                    string PartName = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][5])))
                                    {
                                        try
                                        {
                                            PartName = @dt.Rows[i][5].ToString();
                                            //PartName = RemoveSpecialCharacters(PartName);
                                            PartName = PartName.Replace("'", "");
                                            PartName = PartName.Replace("\"", "");
                                            if (PartName.Length > 0)
                                            {
                                                PartName = PartName.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string type = null;
                                    type = Convert.ToString(dt.Rows[i][6]);
                                    int TargetQty = 0;
                                    string TargQty = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][7])))
                                    {
                                        // //MessageBox.Show("T : " + dt.Rows[i][8].ToString());
                                        try
                                        {
                                            TargQty = Convert.ToString(dt.Rows[i][7]);
                                            TargQty = TargQty.Replace(",", "");
                                            string[] inta = TargQty.Trim().Split('.');
                                            TargetQty = Convert.ToInt32(inta[0]);
                                        }
                                        catch (Exception e)
                                        {
                                            IntoFile("" + e);
                                            continue;
                                        }

                                    }
                                    else
                                    {
                                        IntoFile("TargetQty Empty for" + WorkOrder + " OpNo:" + OpNo);
                                        continue;
                                    }
                                    if (TargetQty == 0)
                                    {
                                        IntoFile("TargetQty 0 for" + WorkOrder + " OpNo:" + OpNo);
                                        continue;
                                    }

                                    string madDate = null;
                                    if (!String.IsNullOrEmpty(Convert.ToString(dt.Rows[i][8])))
                                    {
                                        ////MessageBox.Show("Col 9. Row : " + i + " " + dt.Rows[i][9].ToString());
                                        try
                                        {
                                            string[] words = dt.Rows[i][8].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                // ////MessageBox.Show("INtermeditate : " + IntermediateMadDate);
                                                madDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }

                                    //MessageBox.Show(WorkCenter + " " + WorkOrder + " " + OpNo + " " + OperationDesc + " " + partNo + " " + PartName + " " + type + " " + TargQty + " " + madDate);

                                    string madDateInd = null;
                                    try
                                    {
                                        madDateInd = Convert.ToString(dt.Rows[i][9]);
                                        if (madDateInd.Length > 0)
                                        {
                                            madDateInd = madDateInd.Trim();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        //continue;
                                        IntoFile(e.ToString());
                                    }


                                    string preopnEndDate = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][10].ToString()))
                                    {
                                        try
                                        {
                                            //MessageBox.Show("Col 11. Row : " + i + " " + dt.Rows[i][10].ToString());
                                            string[] words = dt.Rows[i][10].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                preopnEndDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");

                                            }
                                            if (preopnEndDate.Length > 0)
                                            {
                                                preopnEndDate = preopnEndDate.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string DaysAgeing = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][11].ToString()))
                                    {
                                        try
                                        {
                                            //DaysAgeing = Convert.ToString(dt.Rows[i][11]) != null ? Convert.ToString(dt.Rows[i][11]).Trim() : Convert.ToString(dt.Rows[i][11]);
                                            DaysAgeing = Convert.ToString(dt.Rows[i][11]);
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                        if (DaysAgeing.Length > 0)
                                        {
                                            DaysAgeing = DaysAgeing.Trim();
                                        }

                                    }
                                    string Project = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][12].ToString()))
                                    {
                                        try
                                        {
                                            Project = dt.Rows[i][12].ToString();
                                            if (Project.Length > 0)
                                            {
                                                Project = Project.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }

                                    string dueDate = "";
                                    if (!String.IsNullOrEmpty(dt.Rows[i][13].ToString()))
                                    {
                                        try
                                        {
                                            ////MessageBox.Show("Col 14. Row : " + i + " " + dt.Rows[i][14].ToString());
                                            string[] words = dt.Rows[i][13].ToString().Split('.');
                                            if (words.Length > 1)
                                            {
                                                string IntermediateMadDate = words[2] + "-" + words[1] + "-" + words[0];
                                                ////MessageBox.Show("INtermeditate : " + IntermediateMadDate);
                                                dueDate = Convert.ToDateTime(IntermediateMadDate).Date.ToString("yyyy-MM-dd");
                                            }
                                            if (dueDate.Length > 0)
                                            {
                                                dueDate = dueDate.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    flagrush = null;
                                    try
                                    {
                                        flagrush = Convert.ToString(dt.Rows[i][14]);
                                        if (flagrush.Length > 0)
                                        {
                                            flagrush = flagrush.Trim();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        //continue;
                                        IntoFile(e.ToString());
                                    }

                                    string OpOnHold = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][15].ToString()))
                                    {
                                        try
                                        {
                                            OpOnHold = dt.Rows[i][15].ToString();
                                            if (OpOnHold.Length > 0)
                                            {
                                                OpOnHold = OpOnHold.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }
                                    string OpOnHoldReason = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][16].ToString()))
                                    {
                                        try
                                        {
                                            OpOnHoldReason = dt.Rows[i][16].ToString();
                                            if (OpOnHoldReason.Length > 0)
                                            {
                                                OpOnHoldReason = OpOnHoldReason.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }

                                    }
                                    string SplitWO = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[i][17].ToString()))
                                    {
                                        try
                                        {
                                            SplitWO = dt.Rows[i][17].ToString();
                                            if (SplitWO.Length > 0)
                                            {
                                                SplitWO = SplitWO.Trim();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }
                                    if ((!String.IsNullOrEmpty(WorkOrder)) && (!String.IsNullOrEmpty(OpNo)) && (!String.IsNullOrEmpty(partNo)))
                                    {
                                        try
                                        {
                                            using (MsqlConnection mc1 = new MsqlConnection())
                                            {
                                                mc1.open();
                                                SqlCommand cmd2 = null;
                                                //MessageBox.Show("WO & OPNo & PartNo Values: " + WorkOrder + " " + OpNo + " " + partNo);
                                                //cmd2 = new MySqlCommand("INSERT INTO tblmachinedetails (InsertedBy,InsertedOn,IsDeleted,MachineType, MachineInvNo, IPAddress, ControllerType,MachineModel,MachineMake,ModelType,MachineDispName,IsParameters,ShopNo,IsPCB) VALUES( '" + dat + "'," + 0 + "," + 2 + ",'" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1].ToString() + "','" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','" + ds.Tables[0].Rows[i][5].ToString() + "','" + ds.Tables[0].Rows[i][6].ToString() + "','" + ds.Tables[0].Rows[i][7].ToString() + "','" + ds.Tables[0].Rows[i][8].ToString() + "','" + ds.Tables[0].Rows[i][9].ToString() + "')", mc1.msqlConnection);
                                                cmd2 = new SqlCommand("INSERT INTO "+ MsqlConnection.DBSchemaName+".tblddl(WorkCenter,WorkOrder,OperationNo,OperationDesc,MaterialDesc,PartName,Type,TargetQty,MADDate,MADDateInd,PreOpnEndDate,DaysAgeing,Project,DueDate,FlagRushInd,OperationsOnHold,ReasonForHold,SplitWO,InsertedOn,CorrectedDate)VALUES('" + WorkCenter + "','" + WorkOrder + "','" + OpNo + "','" + OperationDesc + "','" + partNo + "','" + PartName + "','" + type + "','" + TargetQty + "','" + madDate + "','" + madDateInd + "','" + preopnEndDate + "','" + DaysAgeing + "','" + Project + "','" + dueDate + "','" + flagrush + "','" + OpOnHold + "','" + OpOnHoldReason + "','" + SplitWO + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" + CorrectedDate + "');", mc1.msqlConnection);
                                                cmd2.ExecuteNonQuery();
                                                mc1.close();
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            //continue;
                                            IntoFile(e.ToString());
                                        }
                                    }

                                }
                            }
                            catch (Exception e)
                            {
                                IntoFile(e.ToString());
                            }
                           
                        }
                        catch (Exception exd)
                        {
                            IntoFile(exd.ToString());
                        }
                    }

                }

               // CheckDDLStatus(1);
            }
            catch (Exception exc)
            {
                IntoFile("HandleDDLFile Section: " + exc);
            }
        }

        public static string RemoveSpecialCharacters(string value)
        {
            char[] specialCharacters = { '\'', '\"', '`' };
            return new String(value.Except(specialCharacters).ToArray());
        }

        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            DataTable tbl = new DataTable();
            try
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {

                    try
                    {
                        using (var stream = System.IO.File.OpenRead(path))
                        {
                            pck.Load(stream);
                        }
                        var ws = pck.Workbook.Worksheets.First();

                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                        }
                        var startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            //18 Columns in out excel fixed.
                            var wsRow = ws.Cells[rowNum, 1, rowNum, 18];
                            DataRow row = tbl.Rows.Add();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        using (Form1 obj = new Form1())
                        {
                            obj.IntoFile("Reading excel Data " + e);
                        }
                    }

                }
            }
            catch (Exception exc)
            {
                Form1 f1 = new Form1();
                f1.IntoFile("GetDataTableFromExcel section: " + exc);
            }
            return tbl;
        }

        void fileSystemWatcher1_Changed(object sender, FileSystemEventArgs e)
        {
            // Occurs when the contents of the file change.
            //MessageBox.Show(string.Format("Changed: {0} {1}", e.FullPath, e.ChangeType));
        }

        void fileSystemWatcher1_Deleted(object sender, FileSystemEventArgs e)
        {
            // FullPath is the location of where the file used to be.
            //MessageBox.Show(string.Format("Deleted: {0} {1}", e.FullPath, e.ChangeType));
        }

        void fileSystemWatcher1_Renamed(object sender, RenamedEventArgs e)
        {
            // FullPath is the new file name.
            //MessageBox.Show(string.Format("Renamed: {0} {1}", e.FullPath, e.ChangeType));
        }

        private void Form1_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            System.Diagnostics.Process.Start(Application.StartupPath);
            Application.Restart();
        }

        public void IntoFile(string Msg)
        {
            string path1 = AppDomain.CurrentDomain.BaseDirectory;
            string appPath = Application.StartupPath + @"\FileWatcherLogFile.txt";
            using (StreamWriter writer = new StreamWriter(appPath, true)) //true => Append Text
            {
                writer.WriteLine(System.DateTime.Now + ":  " + Msg);
            }
        }
    }
}
