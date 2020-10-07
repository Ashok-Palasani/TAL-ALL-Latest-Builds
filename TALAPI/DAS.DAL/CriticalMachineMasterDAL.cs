using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.CriticalMachineMasterEntity;

namespace DAS.DAL
{
    public class CriticalMachineMasterDAL : ICriticalMachineMaster
    {

        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(CriticalMachineMasterDAL));

        public CriticalMachineMasterDAL(i_facility_talContext _db)
        {
            db = _db;
        }

        /// <summary>
        /// Get Plants
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetPlants()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.Tblplant
                             where wf.IsDeleted == 0
                             select new
                             {
                                 PlantId = wf.PlantId,
                                 PlantName = wf.PlantName,
                                 PlantDesc = wf.PlantDesc
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
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
        /// Get Shops
        /// </summary>
        /// <param name="plantId"></param>
        /// <returns></returns>
        public CommonResponse1 GetShops(int plantId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.Tblshop
                             where wf.IsDeleted == 0 && wf.PlantId == plantId
                             select new
                             {
                                 ShopId = wf.ShopId,
                                 ShopName = wf.ShopName,
                                 ShopDesc = wf.ShopDesc
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
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
        /// Get Cells
        /// </summary>
        /// <param name="shopId"></param>
        /// <returns></returns>
        public CommonResponse1 GetCells(int shopId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.Tblcell
                             where wf.IsDeleted == 0 && wf.ShopId == shopId
                             select new
                             {
                                 CellId = wf.CellId,
                                 CellName = wf.CellName,
                                 CellDesc = wf.CellDesc
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
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
        /// Get Machines
        /// </summary>
        /// <param name="cellId"></param>
        /// <returns></returns>
        public CommonResponseForCount GetMachines(int cellId)
        {
            CommonResponseForCount obj = new CommonResponseForCount();
            try
            {
                var check = (from wf in db.Tblmachinedetails
                             where wf.CellId == cellId && wf.IsDeleted == 0 && wf.IsCriticalMachine == null
                             select new
                             {
                                 MachineId = wf.MachineId,
                                 MachineInvNo = wf.MachineInvNo,
                                 MachineDispName = wf.MachineDispName,
                                 MachineMake = wf.MachineMake,
                                 MachineModel = wf.MachineModel
                             }).ToList();
                if (check.Count > 0)
                {
                    obj.isStatus = true;
                    obj.response = check;
                    obj.count = check.Count();
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
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
        /// Adding critical Machines
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 AddCriticalMachines(CriticalMachine data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                if (data.checkedMachineIds != "")
                {
                    string[] checkIds = data.checkedMachineIds.Split(',');
                    if (checkIds != null)
                    {
                        foreach (var item in checkIds)
                        {
                            int machineId = Convert.ToInt32(item);
                            var check = db.TblCriticalMachine.Where(m => m.MachineId == machineId).FirstOrDefault();
                            if (check == null)
                            {
                                TblCriticalMachine tblCriticalMachine = new TblCriticalMachine();
                                tblCriticalMachine.MachineId = machineId;
                                tblCriticalMachine.CorrectedDate = DateTime.Now.ToString("yyyy-MM-dd");
                                tblCriticalMachine.InsertedOn = DateTime.Now;
                                tblCriticalMachine.IsCritical = 1;
                                tblCriticalMachine.IsDeleted = 0;
                                db.TblCriticalMachine.Add(tblCriticalMachine);
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = "Added Successfully";

                                var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == machineId).FirstOrDefault();
                                if (dbCheck != null)
                                {
                                    dbCheck.IsCriticalMachine = 1;
                                    db.SaveChanges();
                                }
                            }
                            else
                            {
                                check.IsCritical = 1;
                                check.ModifiedOn = DateTime.Now;
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = "Updated Successfully";

                                var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == machineId).FirstOrDefault();
                                if (dbCheck != null)
                                {
                                    dbCheck.IsCriticalMachine = 1;
                                    db.SaveChanges();
                                }
                            }
                        }
                    }
                }

                if (data.unCheckedMachineIds != "")
                {
                    string[] unCheckIds = data.unCheckedMachineIds.Split(',');
                    if (unCheckIds != null)
                    {
                        foreach (var item1 in unCheckIds)
                        {
                            int machineId1 = Convert.ToInt32(item1);
                            var check = db.TblCriticalMachine.Where(m => m.MachineId == machineId1).FirstOrDefault();
                            if (check != null)
                            {
                                check.IsCritical = 0;
                                check.ModifiedOn = DateTime.Now;
                                db.SaveChanges();
                                obj.isStatus = true;
                                obj.response = "Updated Successfully";
                            }

                            var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == machineId1).FirstOrDefault();
                            if (dbCheck != null)
                            {
                                dbCheck.IsCriticalMachine = null;
                                db.SaveChanges();
                            }
                        }
                    }
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
        /// View Critical Machines in Index Page
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 ViewCriticalMachines()
        {
            CommonResponse1 obj = new CommonResponse1();
            List<ViewCrititcalMachine> ViewCrititcalMachineList = new List<ViewCrititcalMachine>();
            try
            {
                var check = (from wf in db.TblCriticalMachine
                             where wf.IsDeleted == 0 && wf.IsCritical == 1
                             select new
                             {
                                 criticalMachineID = wf.CriticalMachineId,
                                 machineInvNo = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.MachineInvNo).FirstOrDefault(),
                                 machineDispName = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.MachineDispName).FirstOrDefault(),
                                 plantID = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.PlantId).FirstOrDefault(),
                                 shopID = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.ShopId).FirstOrDefault(),
                                 cellID = db.Tblmachinedetails.Where(m => m.MachineId == wf.MachineId).Select(m => m.CellId).FirstOrDefault(),
                                 correctedDate = wf.CorrectedDate
                             }).ToList();
                if(check.Count > 0)
                {
                    foreach(var item in check)
                    {
                        ViewCrititcalMachine viewCrititcalMachine = new ViewCrititcalMachine();
                        viewCrititcalMachine.crititcalMachineId = item.criticalMachineID;
                        viewCrititcalMachine.machineInvNo = item.machineInvNo;
                        viewCrititcalMachine.machineDispName = item.machineDispName;
                        viewCrititcalMachine.plantName = db.Tblplant.Where(m => m.PlantId == item.plantID).Select(m => m.PlantName).FirstOrDefault();
                        viewCrititcalMachine.shopName = db.Tblshop.Where(m => m.ShopId == item.shopID).Select(m => m.ShopName).FirstOrDefault();
                        viewCrititcalMachine.cellName = db.Tblcell.Where(m => m.CellId == item.cellID).Select(m => m.CellName).FirstOrDefault();
                        viewCrititcalMachine.date = item.correctedDate;
                        ViewCrititcalMachineList.Add(viewCrititcalMachine);
                        obj.isStatus = true;
                        obj.response = ViewCrititcalMachineList;
                    }
                }
                else
                {
                    obj.isStatus = false;
                    obj.response = "No Items Found";
                }
            }
            catch(Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// Delete Crititcal Machine
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public CommonResponse1 DeleteCrititcalMachine(int id)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblCriticalMachine.Where(m => m.CriticalMachineId == id).FirstOrDefault();
                if (check != null)
                {
                    check.IsCritical = 0;
                    check.IsDeleted = 1;
                    check.ModifiedOn = DateTime.Now;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Dleted Successfully";
                }

                var dbCheck = db.Tblmachinedetails.Where(m => m.MachineId == check.MachineId).FirstOrDefault();
                if (dbCheck != null)
                {
                    dbCheck.IsCriticalMachine = null;
                    db.SaveChanges();
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
