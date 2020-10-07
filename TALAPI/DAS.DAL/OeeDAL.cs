using DAS.DAL.Resource;
using DAS.DBModels;
using DAS.EntityModels;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using DAS.Interface;
using Microsoft.Extensions.Configuration;
using static DAS.EntityModels.OeeEntity;

namespace DAS.DAL
{
    public class OeeDAL : IOee
    {

        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(SpliDurationDAL));
        public static IConfiguration configuration;

        public OeeDAL(i_facility_talContext _db, IConfiguration _configuration)
        {
            db = _db;
            configuration = _configuration;
        }


        /// <summary>
        /// Add And Update Oee
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 AddAndUpdateOee(OeeDetails data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblOee.Where(m => m.MachineId == data.machineId).FirstOrDefault();
                if (check == null)
                {
                    TblOee tblOee = new TblOee();
                    tblOee.MachineId = data.machineId;
                    tblOee.StdOee = data.stdOee;
                    tblOee.IsDeleted = 0;
                    tblOee.CreatedOn = DateTime.Now;
                    db.TblOee.Add(tblOee);
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = ResourceResponse.AddedSuccessMessage;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "Std Oee already added for this machine";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Update Oee
        /// </summary>
        /// <param name="oeeId"></param>
        /// <param name="stdOee"></param>
        /// <returns></returns>
        public CommonResponse1 UpdateOee(int oeeId, decimal stdOee)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblOee.Where(m => m.OeeId == oeeId).FirstOrDefault();
                if (check != null)
                {
                    check.StdOee = stdOee;
                    check.ModifiedOn = DateTime.Now;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = ResourceResponse.UpdatedSuccessMessage;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Delete Oee
        /// </summary>
        /// <param name="oeeId"></param>
        /// <returns></returns>
        public CommonResponse1 DeleteOee(int oeeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblOee.Where(m => m.OeeId == oeeId).FirstOrDefault();
                if (check != null)
                {
                    check.IsDeleted = 1;
                    check.ModifiedOn = DateTime.Now;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = ResourceResponse.DeletedSuccessMessage;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// View Oee Details
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 ViewOeeDetails()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblOee
                             where wf.IsDeleted == 0
                             select new
                             {
                                 oeeId = wf.OeeId,
                                 machineId = wf.MachineId,
                                 machineName = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.MachineInvNo).FirstOrDefault(),
                                 plantId = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.PlantId).FirstOrDefault(),
                                 shopId = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.ShopId).FirstOrDefault(),
                                 cellId = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.CellId).FirstOrDefault(),
                                 plantName = db.Tblplant.Where(m => m.PlantId == (db.Tblmachinedetails.Where(n => n.MachineId == wf.MachineId)).Select(n => n.PlantId).FirstOrDefault()).Select(m => m.PlantName).FirstOrDefault(),
                                 shopName = db.Tblshop.Where(m => m.ShopId == (db.Tblmachinedetails.Where(n => n.MachineId == wf.MachineId)).Select(n => n.ShopId).FirstOrDefault()).Select(m => m.ShopName).FirstOrDefault(),
                                 cellName = db.Tblcell.Where(m => m.CellId == db.Tblmachinedetails.Where(n => n.MachineId == wf.MachineId).Select(n => m.CellId).FirstOrDefault()).Select(m => m.CellName).FirstOrDefault(),
                                 stdOee = wf.StdOee
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// View  Oe eDetails By Id
        /// </summary>
        /// <param name="oeeId"></param>
        /// <returns></returns>
        public CommonResponse1 ViewOeeDetailsById(int oeeId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblOee
                             where wf.IsDeleted == 0 && wf.OeeId == oeeId
                             select new
                             {
                                 oeeId = wf.OeeId,
                                 machineId = wf.MachineId,
                                 machineName = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.MachineInvNo).FirstOrDefault(),
                                 stdOee = wf.StdOee
                             }).FirstOrDefault();
                if (check != null)
                {
                    obj.isStatus = true;
                    obj.response = check;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }
    }
}
