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
   public class DALEmployee:IEmployee
    {
        public i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALEmployee));

        public DALEmployee(i_facility_talContext _db)
        {
            db = _db;
        }

        public CommonResponse GetEmployee()
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<GetEmp> actlist = new List<GetEmp>();
                var actList = db.TblEmployee.Where(m => m.Isdeleted == 0).ToList();
                if (actList.Count > 0)
                {
                    foreach (var row in actList)
                    {
                        GetEmp objact = new GetEmp();
                       objact.EId = row.Eid;
                        objact.EmpName = row.EmpName;
                        objact.EmpId = (int)row.EmpId;
                        objact.EmpDesignation = row.EmpDesignation;
                        objact.EmpRole = (int)row.EmpRole;
                        int id= (int)row.EmpRole;
                        var plantName = db.Tblplant.Where(m => m.PlantId == row.PlantId).Select(m => m.PlantName).FirstOrDefault();
                        var shopName = db.Tblshop.Where(m => m.ShopId == row.ShopId).Select(m => m.ShopName).FirstOrDefault();
                        var cellName = db.Tblcell.Where(m => m.CellId == row.CellId).Select(m => m.CellName).FirstOrDefault();
                        var RoleName = db.Tblroles.Where(m => m.RoleId == id).Select(m => m.RoleDesc).FirstOrDefault();
                        objact.PlantName = plantName;
                        objact.ShopName = shopName;
                        objact.rollName = RoleName;
                        objact.CellName = cellName;
                        objact.PlantId = (int)row.PlantId;
                        objact.ShopId = (int)row.ShopId;
                        objact.CellId = (int)row.CellId;
                        objact.CantactNo = row.CantactNo;
                        objact.emailId = row.EmailId;
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

        public EntityModel CreateEmployee(EntityEmployee data)
        {
            EntityModel entity = new EntityModel();
            try
            {
                var act = db.TblEmployee.Where(m => m.Isdeleted == 0 && m.Eid == data.EId).FirstOrDefault();
                if (act == null)
                {
                    int cellid = 0;
                    if(data.CellId != 0)
                    {
                        cellid = data.CellId;
                    }
                    TblEmployee tblact = new TblEmployee();
                    tblact.EmpId = data.EmpId;
                    tblact.EmpName = data.EmpName;
                    tblact.EmpDesignation = data.EmpDesignation;
                    tblact.EmpRole = data.EmpRole;
                    tblact.CantactNo = Convert.ToString(data.CantactNo);
                    tblact.EmailId = data.emailId;
                    tblact.PlantId = data.PlantId;
                    tblact.ShopId = data.ShopId;
                    tblact.CellId = cellid;
                    tblact.Isdeleted = 0;
                    tblact.CreatedOn = DateTime.Now;
                    tblact.CreatedBy = 1;
                    db.TblEmployee.Add(tblact);
                    db.SaveChanges();
                    entity.isTrue = true;
                    entity.response = "Item Created Successfully";
                }
                else
                {
                    int cellid = 0;
                    if (data.CellId != 0)
                    {
                        cellid = data.CellId;
                    }
                    act.EmpName = data.EmpName;
                    act.EmpDesignation = data.EmpDesignation;
                    act.EmpId = data.EmpId;
                    act.EmpRole = data.EmpRole;
                    act.CantactNo = Convert.ToString(data.CantactNo);
                    act.EmailId = data.emailId;
                    act.PlantId = data.PlantId;
                    act.ShopId = data.ShopId;
                    act.CellId = cellid;
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

        public CommonResponse EditEmployee(int id)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                List<GetEmployee> actlist = new List<GetEmployee>();
                var actList = db.TblEmployee.Where(m => m.Isdeleted == 0 && m.Eid == id).FirstOrDefault();

                GetEmployee objact = new GetEmployee();
                objact.EmpName = actList.EmpName;
                objact.EmpDesignation = actList.EmpDesignation;
                objact.EmpId = (int)actList.EmpId;
                objact.EmpRole = (int)actList.EmpRole;
                objact.CantactNo = actList.CantactNo;
                objact.emailId = actList.EmailId;
                objact.PlantId = (int)actList.PlantId;
                objact.ShopId = (int)actList.ShopId;
                objact.CellId = (int)actList.CellId;
                actlist.Add(objact);
                comres.isTure = true;
                comres.response = actlist;
            }
            catch (Exception ex)
            {
                comres.isTure = false;
                //comres.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return comres;
        }

        public CommonResponse GetRole()
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                List<Role> actobj = new List<Role>();
                var act = db.Tblroles.Where(m => m.IsDeleted == 0).ToList();
                foreach (var row in act)
                {
                    Role obj1 = new Role();
                    obj1.Roleid = row.RoleId;
                    obj1.RoleDesc = row.RoleDesc;
                    actobj.Add(obj1);
                }
                if (actobj != null)
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

        public CommonResponse DeleteEmployee(int id)
        {
            CommonResponse comres = new CommonResponse();
            try
            {
                var actList = db.TblEmployee.Where(m => m.Isdeleted == 0 && m.Eid == id).FirstOrDefault();
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
    }
}
