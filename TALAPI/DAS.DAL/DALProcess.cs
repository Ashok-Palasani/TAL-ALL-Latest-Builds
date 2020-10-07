using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAS.DAL
{
    public class DALProcess : IProcess
    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALProcess));

        public DALProcess(i_facility_talContext _db)
        {
            db = _db;
        }

        public EntityModel CreateProcess(EntityProcess data)
        {
            EntityModel entity = new EntityModel();
            try
            {
                var act = db.TblProcess.Where(m => m.Isdeleted == 0 && m.ProcessId == data.processId).FirstOrDefault();
                if (act == null)
                {
                    TblProcess tblact = new TblProcess();
                    tblact.ProcessName = data.processName;
                    tblact.ProcessDescc = data.processDesc;
                    tblact.Isdeleted = 0;
                    tblact.Createdon = DateTime.Now;
                    tblact.CreatedBy = 1;
                    db.TblProcess.Add(tblact);
                    db.SaveChanges();
                    entity.isTrue = true;
                    entity.response = "Item Created Successfully";
                }
                else
                {
                    act.ProcessName = data.processName;
                    act.ProcessDescc = data.processDesc;
                    act.ModifiedOn = DateTime.Now;
                    act.ModifiedBy = 1;
                    db.Entry(act).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                    db.SaveChanges();
                    entity.isTrue = true;
                    entity.response = "Item Updated Successfully";
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

        public CommonResponse DeleteProcess(int id)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                var actList = db.TblProcess.Where(m => m.Isdeleted == 0 && m.ProcessId == id).FirstOrDefault();
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

        public CommonResponse GetProcess()
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<EntityProcess> actlist = new List<EntityProcess>();
                var actList = db.TblProcess.Where(m => m.Isdeleted == 0).ToList();
                if (actList.Count > 0)
                {
                    foreach (var row in actList)
                    {
                        EntityProcess objact = new EntityProcess();
                        objact.processId = row.ProcessId;
                        objact.processName = row.ProcessName;
                        objact.processDesc = row.ProcessDescc;
                        actlist.Add(objact);
                    }
                    comres.isTure = true;
                    comres.response = actlist;
                }
                else
                {
                    comres.isTure = false;
                    //comres.response = ResourceResponse.NoItemsFound; ;
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

        public CommonResponse AddUploadedProcessDetails(proclist data)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                foreach (var dataitem in data.processlist)
                {
                    var processdet = db.TblProcess.Where(m => m.ProcessName == dataitem.Name).FirstOrDefault();
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
                if (obj.isTure == true)
                {
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

        //public CommonResponse EditProcess(int id)
        //{
        //    CommonResponse comres = new CommonResponse();
        //    try
        //    {
        //        List<EditProcess> actlist = new List<EditProcess>();
        //        var actList = db.TblProcess.Where(m => m.Isdeleted == 0 && m.ProcessId == id).FirstOrDefault();

        //        EditProcess objact = new EditProcess();
        //        objact.ProcessName = actList.ProcessName;
        //        objact.ProcessDesc = actList.ProcessDescc;
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
    }
}
