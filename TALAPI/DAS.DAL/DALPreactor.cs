using DAS.DAL.Helpers;
using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace DAS.DAL
{
    public class DALPreactor : IPreactorSchedule
    {
        public i_facility_talContext db = new i_facility_talContext();
        private readonly AppSettings appSettings;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALPreactor));

        public DALPreactor(i_facility_talContext _db, IOptions<AppSettings> _appSettings)
        {
            db = _db;
            appSettings = _appSettings.Value;
        }

        public CommonResponse CreatePreactorSchedule(PreactorEntity data)
        {
            CommonResponse entity = new CommonResponse();
            try
            {
                if (data.ScheduleId == 0)
                {
                    Tblpreactorschedule tblact = new Tblpreactorschedule();
                    tblact.StartTime = data.StartTime;
                    tblact.EndTime = data.EndTime;
                    tblact.DomainName = data.DomainName;
                    tblact.IsNetwork = Convert.ToInt32(data.IsNetwork);
                    tblact.IsStart = Convert.ToInt32(data.isstart);
                    tblact.FileFormat = data.fileFormat;
                    tblact.OutputGenerationTime = TimeSpan.Parse(data.OutputGenerationTime);
                    tblact.Password = data.Password;
                    tblact.UserName = data.UserName;
                    tblact.Path = data.Path;
                    tblact.IsDeleted = 0;
                    tblact.InsertedOn = DateTime.Now;
                    tblact.InsertedBy = 1;
                    db.Tblpreactorschedule.Add(tblact);
                    db.SaveChanges();
                    entity.isTure = true;
                    entity.response = "Item Created Successfully";

                }
                else
                {
                    var tblact = db.Tblpreactorschedule.Where(m => m.ScheduleId == data.ScheduleId).FirstOrDefault();
                    if (tblact != null)
                    {
                        tblact.StartTime = data.StartTime;
                        tblact.EndTime = data.EndTime;
                        tblact.DomainName = data.DomainName;
                        tblact.IsNetwork = Convert.ToInt32(data.IsNetwork);
                        tblact.FileFormat = data.fileFormat;
                        tblact.IsStart = Convert.ToInt32(data.isstart);
                        tblact.OutputGenerationTime = TimeSpan.Parse(data.OutputGenerationTime);
                        tblact.Password = data.Password;
                        tblact.UserName = data.UserName;
                        tblact.Path = data.Path;
                        tblact.ModifiedOn = DateTime.Now;
                        tblact.ModifiedBy = 1;
                        db.Entry(tblact).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                        db.SaveChanges();
                        entity.isTure = true;
                        entity.response = "Item Updated Successfully";
                    }
                }
            }
            catch (Exception ex)
            {
                entity.isTure = false;
                entity.response = ex;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        public CommonResponse DeletePreactorSchedule(int id)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                var actList = db.Tblpreactorschedule.Where(m => m.IsDeleted == 0 && m.ScheduleId == id).FirstOrDefault();
                actList.IsDeleted = 1;
                db.Entry(actList).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                db.SaveChanges();
                comres.isTure = true;
                comres.response = "Item Deleted Successfully";
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ex;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        public CommonResponse GetPreactorSchedule()
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<PreactorEntity> actlist = new List<PreactorEntity>();
                var actList = db.Tblpreactorschedule.Where(m => m.IsDeleted == 0).ToList();
                if (actList.Count > 0)
                {
                    foreach (var row in actList)
                    {
                        PreactorEntity objact = new PreactorEntity();
                        objact.ScheduleId = row.ScheduleId;
                        objact.StartTime = row.StartTime;
                        objact.EndTime = row.EndTime;
                        objact.OutputGenerationTime = Convert.ToString(row.OutputGenerationTime);
                        objact.DomainName = row.DomainName;
                        objact.UserName = row.UserName;
                        objact.Password = row.Password;
                        objact.Path = row.Path;
                        if (row.IsStart == 1)
                        {
                            objact.isstart = "Yes";
                        }
                        else
                        {
                            objact.isstart = "No";
                        }
                        if (row.IsNetwork == 1)
                        {
                            objact.IsNetwork = "Yes";
                        }
                        else
                        {
                            objact.IsNetwork = "No";
                        }
                        objact.fileFormat = row.FileFormat;
                        actlist.Add(objact);
                    }
                    comres.isTure = true;
                    comres.response = actlist;
                }
                else
                {
                    comres.isTure = false;
                    comres.response = ResourceResponse.NoItemsFound; ;
                }
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                comres.response = ex;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        // New Code added

        public CommonResponse1 AddUploadedProcessDetails(proclist data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.type == "NewList")
                {
                    foreach (var dataitem in data.processlist)
                    {
                        var processdet = db.TblProcess.Where(m => m.ProcessName == dataitem.Name && m.Isdeleted == 0).FirstOrDefault();
                        if (processdet != null)
                        {
                            processdet.Isdeleted = 1;
                            db.SaveChanges();
                        }
                        else
                        {
                            TblProcess prcobj = new TblProcess();
                            prcobj.ProcessName = dataitem.Name;
                            prcobj.ProcessDescc = dataitem.Description;
                            prcobj.CreatedBy = 1;
                            prcobj.Createdon = DateTime.Now;
                            prcobj.Isdeleted = 0;
                            db.TblProcess.Add(prcobj);
                            db.SaveChanges();
                            obj.isTure = true;
                        }
                    }
                }
                else if (data.type == "OverWrite")
                {
                    //string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //SqlConnection conn = new SqlConnection(connectionstring);

                    //SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblProcess]", conn);
                    //conn.Open();
                    //truncateQuery.ExecuteNonQuery();
                    //conn.Close();

                    foreach (var dataitem in data.processlist)
                    {
                        var processdet = db.TblProcess.Where(m => m.ProcessName == dataitem.Name && m.Isdeleted == 0).FirstOrDefault();
                        if (processdet != null)
                        {
                            obj.isTure = false;
                            obj.response = "Duplicate Entry";
                        }
                        else
                        {
                            TblProcess prcobj = new TblProcess();
                            prcobj.ProcessName = dataitem.Name;
                            prcobj.ProcessDescc = dataitem.Description;
                            prcobj.CreatedBy = 1;
                            prcobj.Createdon = DateTime.Now;
                            prcobj.Isdeleted = 0;
                            db.TblProcess.Add(prcobj);
                            db.SaveChanges();
                            obj.isTure = true;
                        }
                    }
                }
                else if (data.type == "Update")
                {
                    foreach (var dataitem in data.processlist)
                    {
                        var processdet = db.TblProcess.Where(m => m.ProcessName == dataitem.Name && m.Isdeleted == 0).FirstOrDefault();
                        if (processdet != null)
                        {
                            processdet.ProcessName = dataitem.Name;
                            processdet.ProcessDescc = dataitem.Description;
                            db.SaveChanges();
                            obj.isTure = true;
                        }
                        else
                        {
                            TblProcess prcobj = new TblProcess();
                            prcobj.ProcessName = dataitem.Name;
                            prcobj.ProcessDescc = dataitem.Description;
                            prcobj.CreatedBy = 1;
                            prcobj.Createdon = DateTime.Now;
                            prcobj.Isdeleted = 0;
                            db.TblProcess.Add(prcobj);
                            db.SaveChanges();
                            obj.isTure = true;
                        }
                    }
                }

                if (obj.isTure == true)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = "item added successfully";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No item added successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        //public CommonResponse1 AddUploadedActivityDetails(actlist data)
        //{
        //    CommonResponse1 obj = new CommonResponse1();
        //    try
        //    {
        //        if (data.type == "NewList")
        //        {
        //            foreach (var dataitem in data.activitylist)
        //            {
        //                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        //                if (plantid != 0)
        //                {
        //                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        //                    if (shopid != 0)
        //                    {
        //                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        //                        if (cellid != 0)
        //                        {
        //                            var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        //                            if (procid != 0)
        //                            {
        //                                var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.CellId == cellid && m.Isdeleted == 0).FirstOrDefault();
        //                                if (actdet != null)
        //                                {
        //                                    actdet.Isdeleted = 1;
        //                                    db.SaveChanges();

        //                                    Tblactivity prcobj = new Tblactivity();
        //                                    prcobj.ActivityName = dataitem.Name;
        //                                    prcobj.ActivityDescc = dataitem.Description;
        //                                    prcobj.ProcessId = procid;
        //                                    prcobj.CellId = cellid;
        //                                    prcobj.ShopId = shopid;
        //                                    prcobj.PlantId = plantid;
        //                                    prcobj.OptionalAct = dataitem.IsOptional;
        //                                    prcobj.CreatedBy = 1;
        //                                    prcobj.Createdon = DateTime.Now;
        //                                    prcobj.Isdeleted = 0;
        //                                    db.Tblactivity.Add(prcobj);
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                                else
        //                                {
        //                                    Tblactivity prcobj = new Tblactivity();
        //                                    prcobj.ActivityName = dataitem.Name;
        //                                    prcobj.ActivityDescc = dataitem.Description;
        //                                    prcobj.ProcessId = procid;
        //                                    prcobj.CellId = cellid;
        //                                    prcobj.ShopId = shopid;
        //                                    prcobj.PlantId = plantid;
        //                                    prcobj.OptionalAct = dataitem.IsOptional;
        //                                    prcobj.CreatedBy = 1;
        //                                    prcobj.Createdon = DateTime.Now;
        //                                    prcobj.Isdeleted = 0;
        //                                    db.Tblactivity.Add(prcobj);
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        else if (data.type == "Update")
        //        {
        //            foreach (var dataitem in data.activitylist)
        //            {
        //                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        //                if (plantid != 0)
        //                {
        //                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        //                    if (shopid != 0)
        //                    {
        //                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        //                        if (cellid != 0)
        //                        {
        //                            var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        //                            if (procid != 0)
        //                            {
        //                                var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.CellId == cellid && m.Isdeleted == 0).FirstOrDefault();
        //                                if (actdet != null)
        //                                {
        //                                    actdet.ActivityName = dataitem.Name;
        //                                    actdet.ActivityDescc = dataitem.Description;
        //                                    actdet.CellId = cellid;
        //                                    actdet.ProcessId = procid;
        //                                    actdet.PlantId = plantid;
        //                                    actdet.ShopId = shopid;
        //                                    actdet.OptionalAct = dataitem.IsOptional;
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                                else
        //                                {
        //                                    Tblactivity prcobj = new Tblactivity();
        //                                    prcobj.ActivityName = dataitem.Name;
        //                                    prcobj.ActivityDescc = dataitem.Description;
        //                                    prcobj.ProcessId = procid;
        //                                    prcobj.CellId = cellid;
        //                                    prcobj.ShopId = shopid;
        //                                    prcobj.PlantId = plantid;
        //                                    prcobj.OptionalAct = dataitem.IsOptional;
        //                                    prcobj.CreatedBy = 1;
        //                                    prcobj.Createdon = DateTime.Now;
        //                                    prcobj.Isdeleted = 0;
        //                                    db.Tblactivity.Add(prcobj);
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        else if (data.type == "OverWrite")
        //        {
        //            //string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
        //            //string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
        //            //SqlConnection conn = new SqlConnection(connectionstring);

        //            //SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblActivity]", conn);
        //            //conn.Open();
        //            //truncateQuery.ExecuteNonQuery();
        //            //conn.Close();

        //            foreach (var dataitem in data.activitylist)
        //            {
        //                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        //                if (plantid != 0)
        //                {
        //                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        //                    if (shopid != 0)
        //                    {
        //                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        //                        if (cellid != 0)
        //                        {
        //                            var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        //                            if (procid != 0)
        //                            {
        //                                var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.CellId == cellid && m.Isdeleted == 0).FirstOrDefault();
        //                                if (actdet != null)
        //                                {
        //                                    obj.isTure = false;
        //                                    obj.response = "Duplicate Entry";
        //                                }
        //                                else
        //                                {
        //                                    Tblactivity prcobj = new Tblactivity();
        //                                    prcobj.ActivityName = dataitem.Name;
        //                                    prcobj.ActivityDescc = dataitem.Description;
        //                                    prcobj.ProcessId = procid;
        //                                    prcobj.CellId = cellid;
        //                                    prcobj.ShopId = shopid;
        //                                    prcobj.PlantId = plantid;
        //                                    prcobj.OptionalAct = dataitem.IsOptional;
        //                                    prcobj.CreatedBy = 1;
        //                                    prcobj.Createdon = DateTime.Now;
        //                                    prcobj.Isdeleted = 0;
        //                                    db.Tblactivity.Add(prcobj);
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        if (obj.isTure == true)
        //        {
        //            obj.isStatus = true;
        //            obj.isTure = true;
        //            obj.response = "item added successfully";
        //        }
        //        else
        //        {
        //            obj.isTure = false;
        //            obj.response = "No item added successfully";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //        obj.isTure = false;
        //    }
        //    return obj;
        //}

        ////public CommonResponse1 AddUploadedActivityDetails(actlist data)
        ////{
        ////    CommonResponse1 obj = new CommonResponse1();
        ////    try
        ////    {
        ////        if (data.type == "NewList")
        ////        {
        ////            string connectionString1 = Path.Combine(appSettings.DefaultConnection);
        ////            using (SqlConnection sqlConn = new SqlConnection(connectionString1))
        ////            {
        ////                string sql = "Truncate Table [i_facility_PreactorTSAL].[dbo].[tblactivity]";
        ////                using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
        ////                {
        ////                    sqlConn.Open();
        ////                    sqlCmd.ExecuteNonQuery();
        ////                }
        ////            }

        ////            foreach (var dataitem in data.activitylist)
        ////            {
        ////                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        ////                if (plantid != 0)
        ////                {
        ////                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        ////                    if (shopid != 0)
        ////                    {
        ////                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        ////                        if (cellid != 0)
        ////                        {
        ////                            var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
        ////                            if (machineId != 0)
        ////                            {
        ////                                var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        ////                                if (procid != 0)
        ////                                {
        ////                                    var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
        ////                                    if (actdet == null)
        ////                                    {
        ////                                        //actdet.Isdeleted = 1;
        ////                                        //db.SaveChanges();

        ////                                        Tblactivity prcobj = new Tblactivity();
        ////                                        prcobj.ActivityName = dataitem.Name;
        ////                                        prcobj.ActivityDescc = dataitem.Description;
        ////                                        prcobj.ProcessId = procid;
        ////                                        prcobj.MachineId = machineId;
        ////                                        prcobj.CellId = cellid;
        ////                                        prcobj.ShopId = shopid;
        ////                                        prcobj.PlantId = plantid;
        ////                                        prcobj.OptionalAct = dataitem.IsOptional;
        ////                                        prcobj.CreatedBy = 1;
        ////                                        prcobj.Createdon = DateTime.Now;
        ////                                        prcobj.Isdeleted = 0;
        ////                                        db.Tblactivity.Add(prcobj);
        ////                                        db.SaveChanges();
        ////                                        obj.isTure = true;
        ////                                    }
        ////                                    //else
        ////                                    //{
        ////                                    //    Tblactivity prcobj = new Tblactivity();
        ////                                    //    prcobj.ActivityName = dataitem.Name;
        ////                                    //    prcobj.ActivityDescc = dataitem.Description;
        ////                                    //    prcobj.ProcessId = procid;
        ////                                    //    prcobj.MachineId = machineId;
        ////                                    //    prcobj.CellId = cellid;
        ////                                    //    prcobj.ShopId = shopid;
        ////                                    //    prcobj.PlantId = plantid;
        ////                                    //    prcobj.OptionalAct = dataitem.IsOptional;
        ////                                    //    prcobj.CreatedBy = 1;
        ////                                    //    prcobj.Createdon = DateTime.Now;
        ////                                    //    prcobj.Isdeleted = 0;
        ////                                    //    db.Tblactivity.Add(prcobj);
        ////                                    //    db.SaveChanges();
        ////                                    //    obj.isTure = true;
        ////                                    //}
        ////                                }
        ////                            }
        ////                        }
        ////                    }
        ////                }
        ////            }
        ////        }

        ////        else if (data.type == "Update")
        ////        {
        ////            foreach (var dataitem in data.activitylist)
        ////            {
        ////                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        ////                if (plantid != 0)
        ////                {
        ////                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        ////                    if (shopid != 0)
        ////                    {
        ////                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        ////                        if (cellid != 0)
        ////                        {
        ////                            var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
        ////                            if (machineId != 0)
        ////                            {
        ////                                var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        ////                                if (procid != 0)
        ////                                {
        ////                                    var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
        ////                                    if (actdet != null)
        ////                                    {
        ////                                        actdet.ActivityName = dataitem.Name;
        ////                                        actdet.ActivityDescc = dataitem.Description;
        ////                                        actdet.MachineId = machineId;
        ////                                        actdet.CellId = cellid;
        ////                                        actdet.ProcessId = procid;
        ////                                        actdet.PlantId = plantid;
        ////                                        actdet.ShopId = shopid;
        ////                                        actdet.OptionalAct = dataitem.IsOptional;
        ////                                        db.SaveChanges();
        ////                                        obj.isTure = true;
        ////                                    }
        ////                                    else
        ////                                    {
        ////                                        Tblactivity prcobj = new Tblactivity();
        ////                                        prcobj.ActivityName = dataitem.Name;
        ////                                        prcobj.ActivityDescc = dataitem.Description;
        ////                                        prcobj.ProcessId = procid;
        ////                                        prcobj.MachineId = machineId;
        ////                                        prcobj.CellId = cellid;
        ////                                        prcobj.ShopId = shopid;
        ////                                        prcobj.PlantId = plantid;
        ////                                        prcobj.OptionalAct = dataitem.IsOptional;
        ////                                        prcobj.CreatedBy = 1;
        ////                                        prcobj.Createdon = DateTime.Now;
        ////                                        prcobj.Isdeleted = 0;
        ////                                        db.Tblactivity.Add(prcobj);
        ////                                        db.SaveChanges();
        ////                                        obj.isTure = true;
        ////                                    }
        ////                                }
        ////                            }
        ////                        }
        ////                    }
        ////                }
        ////            }
        ////        }

        ////        else if (data.type == "OverWrite")
        ////        {
        ////            //string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
        ////            //string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
        ////            //SqlConnection conn = new SqlConnection(connectionstring);

        ////            //SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblActivity]", conn);
        ////            //conn.Open();
        ////            //truncateQuery.ExecuteNonQuery();
        ////            //conn.Close();

        ////            foreach (var dataitem in data.activitylist)
        ////            {
        ////                var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        ////                if (plantid != 0)
        ////                {
        ////                    var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        ////                    if (shopid != 0)
        ////                    {
        ////                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        ////                        if (cellid != 0)
        ////                        {
        ////                            var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
        ////                            if (machineId != 0)
        ////                            {
        ////                                var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        ////                                if (procid != 0)
        ////                                {
        ////                                    var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
        ////                                    if (actdet != null)
        ////                                    {
        ////                                        obj.isTure = false;
        ////                                        obj.response = "Duplicate Entry";
        ////                                    }
        ////                                    else
        ////                                    {
        ////                                        Tblactivity prcobj = new Tblactivity();
        ////                                        prcobj.ActivityName = dataitem.Name;
        ////                                        prcobj.ActivityDescc = dataitem.Description;
        ////                                        prcobj.ProcessId = procid;
        ////                                        prcobj.MachineId = machineId;
        ////                                        prcobj.CellId = cellid;
        ////                                        prcobj.ShopId = shopid;
        ////                                        prcobj.PlantId = plantid;
        ////                                        prcobj.OptionalAct = dataitem.IsOptional;
        ////                                        prcobj.CreatedBy = 1;
        ////                                        prcobj.Createdon = DateTime.Now;
        ////                                        prcobj.Isdeleted = 0;
        ////                                        db.Tblactivity.Add(prcobj);
        ////                                        db.SaveChanges();
        ////                                        obj.isTure = true;
        ////                                    }
        ////                                }
        ////                            }
        ////                        }
        ////                    }
        ////                }
        ////            }
        ////        }

        ////        if (obj.isTure == true)
        ////        {
        ////            obj.isStatus = true;
        ////            obj.isTure = true;
        ////            obj.response = "item added successfully";
        ////        }
        ////        else
        ////        {
        ////            obj.isTure = false;
        ////            obj.response = "No item added successfully";
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        ////        obj.isTure = false;
        ////    }
        ////    return obj;
        ////}   //Added Machinid to upload the activity details

        public CommonResponse1 AddUploadedActivityDetails(actlist data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.type == "NewList")
                {
                    string connectionString1 = Path.Combine(appSettings.DefaultConnection);
                    using (SqlConnection sqlConn = new SqlConnection(connectionString1))
                    {
                        string sql = "Truncate Table [i_facility_tal].[dbo].[tblactivity]";
                        using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
                        {
                            sqlConn.Open();
                            sqlCmd.ExecuteNonQuery();
                        }
                    }

                    foreach (var dataitem in data.activitylist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                if (cellid != 0)
                                {
                                    var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
                                    if (machineId != 0)
                                    {
                                        var procid = db.TblProcess.Where(m => m.ProcessDescc == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
                                        if (procid != 0)
                                        {
                                            var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
                                            if (actdet == null)
                                            {
                                                //actdet.Isdeleted = 1;
                                                //db.SaveChanges();

                                                Tblactivity prcobj = new Tblactivity();
                                                prcobj.ActivityName = dataitem.Name;
                                                prcobj.ActivityDescc = dataitem.Description;
                                                prcobj.ProcessId = procid;
                                                prcobj.MachineId = machineId;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.OptionalAct = dataitem.IsOptional;
                                                prcobj.CreatedBy = 1;
                                                prcobj.Createdon = DateTime.Now;
                                                prcobj.Isdeleted = 0;
                                                db.Tblactivity.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                            //else
                                            //{
                                            //    Tblactivity prcobj = new Tblactivity();
                                            //    prcobj.ActivityName = dataitem.Name;
                                            //    prcobj.ActivityDescc = dataitem.Description;
                                            //    prcobj.ProcessId = procid;
                                            //    prcobj.MachineId = machineId;
                                            //    prcobj.CellId = cellid;
                                            //    prcobj.ShopId = shopid;
                                            //    prcobj.PlantId = plantid;
                                            //    prcobj.OptionalAct = dataitem.IsOptional;
                                            //    prcobj.CreatedBy = 1;
                                            //    prcobj.Createdon = DateTime.Now;
                                            //    prcobj.Isdeleted = 0;
                                            //    db.Tblactivity.Add(prcobj);
                                            //    db.SaveChanges();
                                            //    obj.isTure = true;
                                            //}
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                else if (data.type == "Update")
                {
                    foreach (var dataitem in data.activitylist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                if (cellid != 0)
                                {
                                    var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
                                    if (machineId != 0)
                                    {
                                        var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
                                        if (procid != 0)
                                        {
                                            var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
                                            if (actdet != null)
                                            {
                                                actdet.ActivityName = dataitem.Name;
                                                actdet.ActivityDescc = dataitem.Description;
                                                actdet.MachineId = machineId;
                                                actdet.CellId = cellid;
                                                actdet.ProcessId = procid;
                                                actdet.PlantId = plantid;
                                                actdet.ShopId = shopid;
                                                actdet.OptionalAct = dataitem.IsOptional;
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                            else
                                            {
                                                Tblactivity prcobj = new Tblactivity();
                                                prcobj.ActivityName = dataitem.Name;
                                                prcobj.ActivityDescc = dataitem.Description;
                                                prcobj.ProcessId = procid;
                                                prcobj.MachineId = machineId;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.OptionalAct = dataitem.IsOptional;
                                                prcobj.CreatedBy = 1;
                                                prcobj.Createdon = DateTime.Now;
                                                prcobj.Isdeleted = 0;
                                                db.Tblactivity.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                else if (data.type == "OverWrite")
                {
                    //string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //SqlConnection conn = new SqlConnection(connectionstring);

                    //SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblActivity]", conn);
                    //conn.Open();
                    //truncateQuery.ExecuteNonQuery();
                    //conn.Close();

                    foreach (var dataitem in data.activitylist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                if (cellid != 0)
                                {
                                    var machineId = db.Tblmachinedetails.Where(m => m.MachineInvNo == dataitem.machineName).Select(m => m.MachineId).FirstOrDefault();
                                    if (machineId != 0)
                                    {
                                        var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
                                        if (procid != 0)
                                        {
                                            var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.MachineId == machineId && m.Isdeleted == 0).FirstOrDefault();
                                            if (actdet != null)
                                            {
                                                obj.isTure = false;
                                                obj.response = "Duplicate Entry";
                                            }
                                            else
                                            {
                                                Tblactivity prcobj = new Tblactivity();
                                                prcobj.ActivityName = dataitem.Name;
                                                prcobj.ActivityDescc = dataitem.Description;
                                                prcobj.ProcessId = procid;
                                                prcobj.MachineId = machineId;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.OptionalAct = dataitem.IsOptional;
                                                prcobj.CreatedBy = 1;
                                                prcobj.Createdon = DateTime.Now;
                                                prcobj.Isdeleted = 0;
                                                db.Tblactivity.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (obj.isTure == true)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = "item added successfully";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No item added successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponse1 AddUploadedApprovedMasterDetails(tcflist data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.type == "NewList")
                {
                    foreach (var dataitem in data.approvedlist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                if (dataitem.cellName != "")
                                {
                                    var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                    if (cellid != 0)
                                    {
                                        var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                        if (moduleid != 0)
                                        {
                                            var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.CellId == cellid && m.TcfModuleId == moduleid).FirstOrDefault();
                                            if (actdet != null)
                                            {
                                                actdet.IsDeleted = 0;
                                                db.SaveChanges();
                                            }
                                            else
                                            {
                                                TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                                prcobj.TcfModuleId = moduleid;
                                                prcobj.FirstApproverCcList = dataitem.firstCCList;
                                                prcobj.FirstApproverToList = dataitem.firstToList;
                                                prcobj.SecondApproverCcList = dataitem.secondCCList;
                                                prcobj.SecondApproverToList = dataitem.secondToList;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.CreatedBy = 1;
                                                prcobj.CreatedOn = DateTime.Now;
                                                prcobj.IsDeleted = 0;
                                                db.TblTcfApprovedMaster.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                    if (moduleid != 0)
                                    {
                                        var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.TcfModuleId == moduleid && m.IsDeleted == 0).FirstOrDefault();
                                        if (actdet != null)
                                        {
                                            actdet.IsDeleted = 0;
                                            db.SaveChanges();
                                        }
                                        else
                                        {
                                            TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                            prcobj.TcfModuleId = moduleid;
                                            prcobj.FirstApproverCcList = dataitem.firstCCList;
                                            prcobj.FirstApproverToList = dataitem.firstToList;
                                            prcobj.SecondApproverCcList = dataitem.secondCCList;
                                            prcobj.SecondApproverToList = dataitem.secondToList;
                                            prcobj.CellId = null;
                                            prcobj.ShopId = shopid;
                                            prcobj.PlantId = plantid;
                                            prcobj.CreatedBy = 1;
                                            prcobj.CreatedOn = DateTime.Now;
                                            prcobj.IsDeleted = 0;
                                            db.TblTcfApprovedMaster.Add(prcobj);
                                            db.SaveChanges();
                                            obj.isTure = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                else if (data.type == "Update")
                {
                    foreach (var dataitem in data.approvedlist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                if (dataitem.cellName != "")
                                {
                                    var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                    if (cellid != 0)
                                    {
                                        var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                        if (moduleid != 0)
                                        {
                                            var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.CellId == cellid && m.TcfModuleId == moduleid && m.IsDeleted == 0).FirstOrDefault();
                                            if (actdet != null)
                                            {
                                                actdet.PlantId = plantid;
                                                actdet.ShopId = shopid;
                                                actdet.CellId = cellid;
                                                actdet.TcfModuleId = moduleid;
                                                actdet.FirstApproverCcList = dataitem.firstCCList;
                                                actdet.FirstApproverToList = dataitem.firstToList;
                                                actdet.SecondApproverCcList = dataitem.secondCCList;
                                                actdet.SecondApproverToList = dataitem.secondToList;
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                            else
                                            {
                                                TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                                prcobj.TcfModuleId = moduleid;
                                                prcobj.FirstApproverCcList = dataitem.firstCCList;
                                                prcobj.FirstApproverToList = dataitem.firstToList;
                                                prcobj.SecondApproverCcList = dataitem.secondCCList;
                                                prcobj.SecondApproverToList = dataitem.secondToList;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.CreatedBy = 1;
                                                prcobj.CreatedOn = DateTime.Now;
                                                prcobj.IsDeleted = 0;
                                                db.TblTcfApprovedMaster.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                    if (moduleid != 0)
                                    {
                                        var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.TcfModuleId == moduleid && m.IsDeleted == 0).FirstOrDefault();
                                        if (actdet != null)
                                        {
                                            actdet.PlantId = plantid;
                                            actdet.ShopId = shopid;
                                            actdet.CellId = null;
                                            actdet.TcfModuleId = moduleid;
                                            actdet.FirstApproverCcList = dataitem.firstCCList;
                                            actdet.FirstApproverToList = dataitem.firstToList;
                                            actdet.SecondApproverCcList = dataitem.secondCCList;
                                            actdet.SecondApproverToList = dataitem.secondToList;
                                            db.SaveChanges();
                                            obj.isTure = true;
                                        }
                                        else
                                        {
                                            TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                            prcobj.TcfModuleId = moduleid;
                                            prcobj.FirstApproverCcList = dataitem.firstCCList;
                                            prcobj.FirstApproverToList = dataitem.firstToList;
                                            prcobj.SecondApproverCcList = dataitem.secondCCList;
                                            prcobj.SecondApproverToList = dataitem.secondToList;
                                            prcobj.CellId = null;
                                            prcobj.ShopId = shopid;
                                            prcobj.PlantId = plantid;
                                            prcobj.CreatedBy = 1;
                                            prcobj.CreatedOn = DateTime.Now;
                                            prcobj.IsDeleted = 0;
                                            db.TblTcfApprovedMaster.Add(prcobj);
                                            db.SaveChanges();
                                            obj.isTure = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                else if (data.type == "OverWrite")
                {
                    //string connectionstring = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //string dbName = configuration.GetSection("MySetting").GetSection("DbConnection").Value;
                    //SqlConnection conn = new SqlConnection(connectionstring);

                    //SqlCommand truncateQuery = new SqlCommand("truncate table " + dbName + ".[tblActivity]", conn);
                    //conn.Open();
                    //truncateQuery.ExecuteNonQuery();
                    //conn.Close();

                    foreach (var dataitem in data.approvedlist)
                    {
                        var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
                        if (plantid != 0)
                        {
                            var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
                            if (shopid != 0)
                            {
                                if (dataitem.cellName != "")
                                {
                                    var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
                                    if (cellid != 0)
                                    {
                                        var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                        if (moduleid != 0)
                                        {
                                            var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.CellId == cellid && m.TcfModuleId == moduleid && m.IsDeleted == 0).FirstOrDefault();
                                            if (actdet != null)
                                            {
                                                obj.isTure = false;
                                                obj.response = "Duplicate Entry";
                                            }
                                            else
                                            {
                                                TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                                prcobj.TcfModuleId = moduleid;
                                                prcobj.FirstApproverCcList = dataitem.firstCCList;
                                                prcobj.FirstApproverToList = dataitem.firstToList;
                                                prcobj.SecondApproverCcList = dataitem.secondCCList;
                                                prcobj.SecondApproverToList = dataitem.secondToList;
                                                prcobj.CellId = cellid;
                                                prcobj.ShopId = shopid;
                                                prcobj.PlantId = plantid;
                                                prcobj.CreatedBy = 1;
                                                prcobj.CreatedOn = DateTime.Now;
                                                prcobj.IsDeleted = 0;
                                                db.TblTcfApprovedMaster.Add(prcobj);
                                                db.SaveChanges();
                                                obj.isTure = true;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
                                    if (moduleid != 0)
                                    {


                                        TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
                                        prcobj.TcfModuleId = moduleid;
                                        prcobj.FirstApproverCcList = dataitem.firstCCList;
                                        prcobj.FirstApproverToList = dataitem.firstToList;
                                        prcobj.SecondApproverCcList = dataitem.secondCCList;
                                        prcobj.SecondApproverToList = dataitem.secondToList;
                                        prcobj.CellId = null;
                                        prcobj.ShopId = shopid;
                                        prcobj.PlantId = plantid;
                                        prcobj.CreatedBy = 1;
                                        prcobj.CreatedOn = DateTime.Now;
                                        prcobj.IsDeleted = 0;
                                        db.TblTcfApprovedMaster.Add(prcobj);
                                        db.SaveChanges();
                                        obj.isTure = true;

                                    }
                                }
                            }
                        }
                    }

                }
                if (obj.isTure == true)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = "item added successfully";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No item added successfully";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponse1 AddUploadedHeattreamentDetails(HTlist data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                foreach (var dataitem in data.htlist)
                {
                    var processdet = db.Tblddl.Where(m => m.WorkOrder == dataitem.WONo && m.OperationNo == dataitem.OPNo && m.PartName == dataitem.PartNo).FirstOrDefault();
                    if (processdet != null)
                    {
                        //processdet.Htno = dataitem.HTNo;
                        db.SaveChanges();
                        obj.isTure = true;
                    }
                }
                if (obj.isTure == true)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = "item added successfully";
                }
                else
                {
                    obj.isTure = false;
                    obj.response = "No item added successfully";
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
        /// Upload Pcp No
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 UploadPcpNo(PcpDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.type == "NewList")
                {
                    string connectionString1 = Path.Combine(appSettings.DefaultConnection);
                    using (SqlConnection sqlConn = new SqlConnection(connectionString1))
                    {
                        string sql = "Truncate Table [i_facility_tsal].[dbo].[tblPcpNo]";
                        using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
                        {
                            sqlConn.Open();
                            sqlCmd.ExecuteNonQuery();
                        }
                    }

                    foreach (var item in data.PCPNoForHeatCycle)
                    {
                        string connectionString = Path.Combine(appSettings.DefaultConnection);
                        using (SqlConnection sqlConn = new SqlConnection(connectionString))
                        {
                            string sql = "InsertPcPNo";
                            using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
                            {
                                sqlCmd.CommandType = CommandType.StoredProcedure;

                                sqlCmd.Parameters.AddWithValue("@partNo", item.partNumber);
                                sqlCmd.Parameters.AddWithValue("@specialProcessInvolved", item.specialProcessInvolved);
                                sqlCmd.Parameters.AddWithValue("@pcpNo", item.textToBeMentionedInRCWithPCPNumbers);
                                sqlCmd.Parameters.AddWithValue("@createdOn", DateTime.Now);
                                sqlConn.Open();
                                sqlCmd.ExecuteNonQuery();

                            }
                        }
                    }

                    foreach (var item in data.PCPNoForHeatCycle)
                    {
                        var check = db.Tblddl.Where(m => m.PartName == item.partNumber).Select(m => m.Ddlid).ToList();
                        if (check.Count > 0)
                        {
                            foreach (var items in check)
                            {
                                var dbCheck = db.Tblddl.Where(m => m.Ddlid == items).FirstOrDefault();
                                if (dbCheck != null)
                                {
                                    dbCheck.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = ResourceResponse.AddedSuccessMessage;
                }
                else if (data.type == "OverWrite")
                {
                    foreach (var item in data.PCPNoForHeatCycle)
                    {
                        var check = db.TblPcpNo.Where(m => m.PartNo == item.partNumber && m.IsDeleted == 0).FirstOrDefault();
                        if (check == null)
                        {
                            TblPcpNo tblPcpNo = new TblPcpNo();
                            tblPcpNo.PartNo = item.partNumber;
                            tblPcpNo.SpecialProcessInvolved = item.specialProcessInvolved;
                            tblPcpNo.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.response = ResourceResponse.AddedSuccessMessage;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.isTure = false;
                            obj.response = "Duplicate Entry";
                        }

                        var check1 = db.Tblddl.Where(m => m.PartName == item.partNumber).Select(m => m.Ddlid).ToList();
                        if (check1.Count > 0)
                        {
                            foreach (var items in check1)
                            {
                                var dbCheck = db.Tblddl.Where(m => m.Ddlid == items).FirstOrDefault();
                                if (dbCheck != null)
                                {
                                    dbCheck.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                }
                else if (data.type == "Update")
                {
                    foreach (var item in data.PCPNoForHeatCycle)
                    {
                        var check = db.TblPcpNo.Where(m => m.PartNo == item.partNumber).FirstOrDefault();
                        if (check != null)
                        {
                            check.SpecialProcessInvolved = item.specialProcessInvolved;
                            check.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                        else
                        {
                            TblPcpNo tblPcpNo = new TblPcpNo();
                            tblPcpNo.PartNo = item.partNumber;
                            tblPcpNo.SpecialProcessInvolved = item.specialProcessInvolved;
                            tblPcpNo.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
                            db.TblPcpNo.Add(tblPcpNo);
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.isTure = true;
                            obj.response = ResourceResponse.AddedSuccessMessage;
                        }

                        var check1 = db.Tblddl.Where(m => m.PartName == item.partNumber).Select(m => m.Ddlid).ToList();
                        if (check1.Count > 0)
                        {
                            foreach (var items in check1)
                            {
                                var dbCheck = db.Tblddl.Where(m => m.Ddlid == items).FirstOrDefault();
                                if (dbCheck != null)
                                {
                                    dbCheck.PcpNo = item.textToBeMentionedInRCWithPCPNumbers;
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
                obj.isTure = false;
            }
            return obj;
        }

        /// <summary>
        /// Get Pcp No
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetPcpNo()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblPcpNo
                             where wf.IsDeleted == 0
                             select new
                             {
                                 partNo = wf.PartNo,
                                 pcpNo = wf.PcpNo,
                                 specialProcessInvolved = wf.SpecialProcessInvolved
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = check;
                }
                else
                {
                    obj.isTure = false;
                    obj.isStatus = false;
                    obj.response = ResourceResponse.NoItemsFound;
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
        /// Upload scrap quantity
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 UploadScrapQuantity(Scrapqnty data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.type == "NewList")
                {
                    string connectionString1 = Path.Combine(appSettings.DefaultConnection);
                    //string connectionString1 = "Server=SRKS-TECH3-PC\\SQLEXPRESS; Database=i_facility_tsal;user id=sa;password=srks4$;";
                    using (SqlConnection sqlConn = new SqlConnection(connectionString1))
                    {
                        string sql = "Truncate Table [i_facility_tal].[dbo].[scrapQty]";
                        using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
                        {
                            sqlConn.Open();
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    foreach (var item in data.ScrapqntyList)
                    {
                        int yield = Convert.ToInt32(item.yield);
                        decimal setup = Convert.ToDecimal(item.setup);
                        decimal mach = Convert.ToDecimal(item.mach);
                        int labour = Convert.ToInt32(item.labour);
                        int cclgcong = Convert.ToInt32(item.ccldConf);
                        int scrapquantity = Convert.ToInt32(item.scrapQuantity);
                        string connectionString = Path.Combine(appSettings.DefaultConnection);
                        DateTime postDate = Convert.ToDateTime(item.PostGDate);
                        string postGDate = postDate.ToString("yyyy-MM-dd");
                        //string connectionString = "Server=SRKS-TECH3-PC\\SQLEXPRESS; Database=i_facility_tsal;user id=sa;password=srks4$;";
                        using (SqlConnection sqlConn = new SqlConnection(connectionString))
                        {
                            string sql = "InsertScrapQuantity";
                            using (SqlCommand sqlCmd = new SqlCommand(sql, sqlConn))
                            {
                                sqlCmd.CommandType = CommandType.StoredProcedure;
                                sqlCmd.Parameters.AddWithValue("@WorkOrder", item.order);
                                sqlCmd.Parameters.AddWithValue("@WorkCenter", item.workCtr);
                                sqlCmd.Parameters.AddWithValue("@Description", item.shortDescription);
                                sqlCmd.Parameters.AddWithValue("@OperationNo", item.opAc);
                                sqlCmd.Parameters.AddWithValue("@PostGDate", postGDate);
                                sqlCmd.Parameters.AddWithValue("@Yield", yield);
                                sqlCmd.Parameters.AddWithValue("@Setup", setup);
                                sqlCmd.Parameters.AddWithValue("@Mach", mach);
                                sqlCmd.Parameters.AddWithValue("@Labour", labour);
                                sqlCmd.Parameters.AddWithValue("@Rev", item.rev);
                                sqlCmd.Parameters.AddWithValue("@CclgCong", cclgcong);
                                sqlCmd.Parameters.AddWithValue("@ScanpQuantity", scrapquantity);
                                sqlCmd.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                                sqlConn.Open();
                                sqlCmd.ExecuteNonQuery();
                            }
                        }
                    }
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = ResourceResponse.AddedSuccessMessage;
                }
                else if (data.type == "OverWrite")
                {
                    foreach (var item in data.ScrapqntyList)
                    {
                        var check = db.ScrapQty.Where(m => m.WorkOder == item.order).FirstOrDefault();
                        if (check == null)
                        {

                            int yield = Convert.ToInt32(item.yield);
                            decimal setup = Convert.ToDecimal(item.setup);
                            decimal mach = Convert.ToDecimal(item.mach);
                            int labour = Convert.ToInt32(item.labour);
                            int cclgcong = Convert.ToInt32(item.ccldConf);
                            int scrapquantity = Convert.ToInt32(item.scrapQuantity);
                            ScrapQty tblScrapqnt = new ScrapQty();
                            tblScrapqnt.WorkOder = item.order;
                            tblScrapqnt.WorkCenter = item.workCtr;
                            tblScrapqnt.Decription = item.shortDescription;
                            tblScrapqnt.OperationNo = item.opAc;
                            tblScrapqnt.PostgDate = item.PostGDate;
                            tblScrapqnt.Yield = yield;
                            tblScrapqnt.Setup = setup;
                            tblScrapqnt.Mach = mach;
                            tblScrapqnt.Labour = labour;
                            tblScrapqnt.CcldConf = cclgcong;
                            tblScrapqnt.NoOfscrapQty = scrapquantity;
                            tblScrapqnt.Rev = item.rev;
                            db.ScrapQty.Add(tblScrapqnt);
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.isTure = true;
                            obj.response = ResourceResponse.AddedSuccessMessage;
                        }
                        else
                        {
                            obj.isStatus = false;
                            obj.isTure = false;
                            obj.response = "Duplicate Entry";
                        }
                    }
                }
                else if (data.type == "Update")
                {
                    foreach (var item in data.ScrapqntyList)
                    {
                        var check = db.ScrapQty.Where(m => m.WorkOder == item.order).FirstOrDefault();
                        if (check != null)
                        {
                            int yield = Convert.ToInt32(item.yield);
                            decimal setup = Convert.ToDecimal(item.setup);
                            decimal mach = Convert.ToDecimal(item.mach);
                            int labour = Convert.ToInt32(item.labour);
                            int cclgcong = Convert.ToInt32(item.ccldConf);
                            int scrapquantity = Convert.ToInt32(item.scrapQuantity);

                            check.WorkOder = item.order;
                            check.WorkCenter = item.workCtr;
                            check.Decription = item.shortDescription;
                            check.OperationNo = item.opAc;
                            check.PostgDate = item.PostGDate;
                            check.Yield = yield;
                            check.Setup = setup;
                            check.Mach = mach;
                            check.Labour = labour;
                            check.CcldConf = cclgcong;
                            check.NoOfscrapQty = scrapquantity;
                            check.Rev = item.rev;
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.isTure = true;
                            obj.response = ResourceResponse.UpdatedSuccessMessage;
                        }
                        else
                        {
                            int yield = Convert.ToInt32(item.yield);
                            decimal setup = Convert.ToDecimal(item.setup);
                            decimal mach = Convert.ToDecimal(item.mach);
                            int labour = Convert.ToInt32(item.labour);
                            int cclgcong = Convert.ToInt32(item.ccldConf);
                            int scrapquantity = Convert.ToInt32(item.scrapQuantity);
                            ScrapQty tblScrapqnt = new ScrapQty();
                            tblScrapqnt.WorkOder = item.order;
                            tblScrapqnt.WorkCenter = item.workCtr;
                            tblScrapqnt.Decription = item.shortDescription;
                            tblScrapqnt.OperationNo = item.opAc;
                            tblScrapqnt.PostgDate = item.PostGDate;
                            tblScrapqnt.Yield = yield;
                            tblScrapqnt.Setup = setup;
                            tblScrapqnt.Mach = mach;
                            tblScrapqnt.Labour = labour;
                            tblScrapqnt.CcldConf = cclgcong;
                            tblScrapqnt.NoOfscrapQty = scrapquantity;
                            tblScrapqnt.Rev = item.rev;
                            db.ScrapQty.Add(tblScrapqnt);
                            db.SaveChanges();
                            obj.isStatus = true;
                            obj.isTure = true;
                            obj.response = ResourceResponse.AddedSuccessMessage;
                        }

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
        /// Get scrap qnty
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetScrapQuantity()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.ScrapQty
                             select new
                             {
                                 workOrder = wf.WorkOder,
                                 workCenter = wf.WorkCenter,
                                 description = wf.Decription,
                                 operationNo = wf.OperationNo,
                                 postGdate = wf.PostgDate,
                                 yield = wf.Yield,
                                 setup = wf.Setup,
                                 mach = wf.Mach,
                                 labour = wf.Labour,
                                 cclgCong = wf.CcldConf,
                                 scanpQuantity = wf.NoOfscrapQty,
                                 rev = wf.Rev
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.isTure = true;
                    obj.response = check;
                }
                else
                {
                    obj.isTure = false;
                    obj.isStatus = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

    }
}
