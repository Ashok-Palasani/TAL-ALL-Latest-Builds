using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAS.DAL
{
   public class DALActivity : IActivity
    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALActivity));

        public DALActivity(i_facility_talContext _db)
        {
            db = _db;
        }
        public CommonResponse GetActivity()
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<EntityActivityForGet> actlist = new List<EntityActivityForGet>();
                var actList = db.Tblactivity.Where(m => m.Isdeleted == 0).ToList();
                if (actList.Count > 0)
                {
                    foreach (var row in actList)
                    {
                        var plantName = db.Tblplant.Where(m => m.PlantId == row.PlantId).Select(m => m.PlantName).FirstOrDefault();
                        var shopName = db.Tblshop.Where(m => m.ShopId == row.ShopId).Select(m => m.ShopName).FirstOrDefault();
                        var cellName = db.Tblcell.Where(m => m.CellId == row.CellId).Select(m => m.CellName).FirstOrDefault();
                        var machineName = db.Tblmachinedetails.Where(m => m.MachineId == row.MachineId).Select(m => m.MachineInvNo).FirstOrDefault();
                        var processName = db.TblProcess.Where(m => m.ProcessId == row.ProcessId).Select(m => m.ProcessName).FirstOrDefault();
                        EntityActivityForGet objact = new EntityActivityForGet();
                        objact.ActivityId = row.ActivityId;
                        objact.ActivityName = row.ActivityName;
                        objact.ActivityDesc = row.ActivityDescc;
                        objact.plantId = (int)row.PlantId;
                        objact.shopId = (int)row.ShopId;
                        objact.cellId = (int)row.CellId;
                        objact.machineId = (int)row.MachineId;
                        objact.machineName = machineName;
                        objact.processId = (int)row.ProcessId;
                        objact.plantName = plantName;
                        objact.shopName = shopName;
                        objact.cellName = cellName;
                        objact.processName = processName;
                        objact.optionalAct = row.OptionalAct;
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
                comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        public CommonResponse GetProcessList()
        {
            CommonResponse retData = new CommonResponse();
            List<GetProcess> listCellEntity = new List<GetProcess>();
            try
            {
                var shopData = db.TblProcess.Where(x => x.Isdeleted == 0).ToList();
                foreach (var row in shopData)
                {
                    GetProcess objCell = new GetProcess();
                    objCell.processId = row.ProcessId;
                    objCell.processName = row.ProcessName;
                    listCellEntity.Add(objCell);
                }
                if (shopData != null)
                {
                    retData.isTure = true;
                    retData.response = listCellEntity;
                }
                else
                {
                    retData.isTure = true;
                    // retData.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                // retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        //Validating the Process whether its already exists for that perticular cell
        public CommonResponse validateprocess(int cellId,int processId)
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                var prodet = db.Tblactivity.Where(m => m.CellId == cellId && m.Isdeleted == 0 && m.ProcessId==processId).FirstOrDefault();
                if(prodet !=null)
                {
                    retData.isTure = false;
                    retData.response = "process is already exists";
                }
                else
                {
                    retData.isTure = true;
                    retData.response = "item not found";
                }
            }
            catch(Exception ex)
            {

            }
            return retData;
        }

        //Validating the Activity whether its already exists for that perticular cell and process
        public CommonResponse validateActivity(validateprocessCell data)
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                var actdet = db.Tblactivity.Where(m => m.CellId == data.cellId && m.Isdeleted == 0 && m.ActivityName == data.activityName && m.ProcessId == data.ProcessId).FirstOrDefault();
                if(actdet!=null)
                {
                    retData.isTure = false;
                    retData.response = "Activity is already exists";
                }
                else
                {
                    retData.isTure = true;
                    retData.response = "item not found";
                }
            }
            catch(Exception ex)
            {

            }
            return retData;
            }

        public EntityModel CreateActivity(EntityActivity data)
        {
            EntityModel entity = new EntityModel();
            try
            {
                var act = db.Tblactivity.Where(m => m.Isdeleted == 0 && m.ActivityName == data.ActivityName).ToList();
                if (act.Count == 0)
                {
                    Tblactivity tblact = new Tblactivity();
                    tblact.ActivityName = data.ActivityName;
                    tblact.ActivityDescc = data.ActivityDesc;
                    tblact.Isdeleted = 0;
                    tblact.Createdon = DateTime.Now;
                    tblact.CreatedBy = 1;
                    tblact.PlantId = Convert.ToInt32(data.plantId);
                    tblact.ShopId = Convert.ToInt32(data.shopId);
                    tblact.CellId = Convert.ToInt32(data.cellId);
                    tblact.MachineId = data.machineId;
                    tblact.ProcessId = Convert.ToInt32(data.processId);
                    tblact.OptionalAct = data.optionalAct;
                    db.Tblactivity.Add(tblact);
                    db.SaveChanges();
                    entity.isTrue = true;
                    entity.response = "Item Created Successfully";
                }
                else
                {
                    entity.isTrue = false;
                    entity.response = "Item Already Exists";
                }
            }
            catch (Exception ex)
            {
                entity.isTrue = false;
                //entity.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return entity;
        }

        //public CommonResponse EditActivityDet(int id)
        //{
        //    CommonResponse comres = new CommonResponse();
        //    try
        //    {
        //        List<EditActivity> actlist = new List<EditActivity>();
        //        var actList = db.Tblactivity.Where(m => m.Isdeleted == 0 && m.ActivityId == id).FirstOrDefault();

        //        EditActivity objact = new EditActivity();
        //        objact.ActivityName = actList.ActivityName;
        //        objact.ActivityDesc = actList.ActivityDescc;
        //        actlist.Add(objact);
        //        comres.isTure = true;
        //        comres.response = actlist;
        //    }
        //    catch (Exception ex)
        //    {
        //        comres.isTure = false;
        //        //comres.response = ResourceResponse.ExceptionMessage;
        //        log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
        //    }
        //    return comres;
        //}

        public CommonResponse UpdateActivity(EntityActivity data)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                var act = db.Tblactivity.Where(m => m.Isdeleted == 0 && m.ActivityName == data.ActivityName && m.ActivityId != data.ActivityId).ToList();
                if (act.Count == 0)
                {
                    var editrecord = db.Tblactivity.Where(m => m.ActivityId == data.ActivityId).FirstOrDefault();
                    editrecord.ActivityName = data.ActivityName;
                    editrecord.ActivityDescc = data.ActivityDesc;
                    editrecord.PlantId = Convert.ToInt32(data.plantId);
                    editrecord.ShopId = Convert.ToInt32(data.shopId);
                    editrecord.CellId = Convert.ToInt32(data.cellId);
                    editrecord.MachineId = data.machineId;
                    editrecord.ProcessId = Convert.ToInt32(data.processId);
                    editrecord.OptionalAct = data.optionalAct;
                    db.Entry(editrecord).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                    db.SaveChanges();
                    comres.isTure = true;
                    comres.response = "Item Updated Successfully"
;
                }
                else
                {
                    comres.isTure = false;
                    comres.response = "item Already Exists";
                }
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                //comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        public CommonResponse DeleteActivity(int id)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                var actList = db.Tblactivity.Where(m => m.Isdeleted == 0 && m.ActivityId == id).FirstOrDefault();
                actList.Isdeleted = 1;
                db.Entry(actList).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                db.SaveChanges();
                comres.isTure = true;
                comres.response = "Item Deleted Successfully";
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                //comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        //public CommonResponse AddUploadedActivityDetails(actlist data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        foreach (var dataitem in data.activitylist)
        //        {
        //            var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        //            if (plantid != 0)
        //            {
        //                var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        //                if (shopid != 0)
        //                {
        //                    var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        //                    if (cellid != 0)
        //                    {
        //                        var procid = db.TblProcess.Where(m => m.ProcessName == dataitem.processName).Select(m => m.ProcessId).FirstOrDefault();
        //                        if (procid != 0)
        //                        {
        //                            var actdet = db.Tblactivity.Where(m => m.ActivityName == dataitem.Name && m.ProcessId == procid && m.CellId == cellid).FirstOrDefault();
        //                            if (actdet != null)
        //                            {
        //                                actdet.ActivityName = dataitem.Name;
        //                                actdet.ActivityDescc = dataitem.Description;
        //                                actdet.CellId = cellid;
        //                                actdet.ProcessId = procid;
        //                                actdet.PlantId = plantid;
        //                                actdet.ShopId = shopid;
        //                                actdet.OptionalAct = dataitem.IsOptional;
        //                                db.SaveChanges();
        //                                obj.isTure = true;
        //                            }
        //                            else
        //                            {
        //                                Tblactivity prcobj = new Tblactivity();
        //                                prcobj.ActivityName = dataitem.Name;
        //                                prcobj.ActivityDescc = dataitem.Description;
        //                                prcobj.ProcessId = procid;
        //                                prcobj.CellId = cellid;
        //                                prcobj.ShopId = shopid;
        //                                prcobj.PlantId = plantid;
        //                                prcobj.OptionalAct = dataitem.IsOptional;
        //                                prcobj.CreatedBy = 1;
        //                                prcobj.Createdon = DateTime.Now;
        //                                prcobj.Isdeleted = 0;
        //                                db.Tblactivity.Add(prcobj);
        //                                db.SaveChanges();
        //                                obj.isTure = true;
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        if (obj.isTure == true)
        //        {
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
    }
}
