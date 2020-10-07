using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.TcfApprovedMasterEntity;

namespace DAS.DAL
{
    public class TcfApprovedMasterDAL : ITcfApprovedMaster
    {
        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(TcfApprovedMasterDAL));

        public TcfApprovedMasterDAL(i_facility_talContext _db)
        {
            db = _db;
        }

        /// <summary>
        /// Add And Edit Tcf Approved Master
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public CommonResponse1 AddAndEditTcfApprovedMaster(AddAndEditTcfMaster data)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblTcfApprovedMaster.Where(m => m.TcfApprovedMasterId == data.TcfApprovedMasterId).FirstOrDefault();
                if (check == null)
                {
                    TblTcfApprovedMaster tblTcfApprovedMaster = new TblTcfApprovedMaster();
                    tblTcfApprovedMaster.TcfModuleId = data.TcfModuleId;
                    tblTcfApprovedMaster.FirstApproverToList = data.FirstApproverToList;
                    tblTcfApprovedMaster.FirstApproverCcList = data.FirstApproverCcList;
                    tblTcfApprovedMaster.SecondApproverToList = data.SecondApproverToList;
                    tblTcfApprovedMaster.SecondApproverCcList = data.SecondApproverCcList;
                    tblTcfApprovedMaster.PlantId = data.PlantId;
                    tblTcfApprovedMaster.ShopId = data.ShopId;
                    tblTcfApprovedMaster.CellId = data.CellId;
                    tblTcfApprovedMaster.CreatedOn = DateTime.Now;
                    tblTcfApprovedMaster.IsDeleted = 0;
                    db.TblTcfApprovedMaster.Add(tblTcfApprovedMaster);
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Added Successfully";
                }
                else
                {
                    check.TcfModuleId = data.TcfModuleId;
                    check.FirstApproverToList = data.FirstApproverToList;
                    check.FirstApproverCcList = data.FirstApproverCcList;
                    check.SecondApproverToList = data.SecondApproverToList;
                    check.SecondApproverCcList = data.SecondApproverCcList;
                    check.PlantId = data.PlantId;
                    check.ShopId = data.ShopId;
                    check.CellId = data.CellId;
                    check.ModifiedOn = DateTime.Now;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Updated Successfully";
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
        /// View Multiple Tcf Approved Master
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 ViewMultipleTcfApprovedMaster()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblTcfApprovedMaster
                             where wf.IsDeleted == 0
                             select new
                             {
                                 TcfApprovedMasterId = wf.TcfApprovedMasterId,
                                 FirstApproverToList = wf.FirstApproverToList,
                                 FirstApproverCcList = wf.FirstApproverCcList,
                                 SecondApproverToList = wf.SecondApproverToList,
                                 SecondApproverCcList = wf.SecondApproverCcList,
                                 plantName = db.Tblplant.Where(m => m.PlantId == wf.PlantId).Select(m => m.PlantName).FirstOrDefault(),
                                 plantId = wf.PlantId,
                                 shopName = db.Tblshop.Where(m => m.ShopId == wf.ShopId).Select(m => m.ShopName).FirstOrDefault(),
                                 shopId = wf.ShopId,
                                 cellName = db.Tblcell.Where(m => m.CellId == wf.CellId).Select(m => m.CellName).FirstOrDefault(),
                                 cellId = wf.CellId,
                                 ModuleName = db.TblTcfModule.Where(m => m.TcfModuleId == wf.TcfModuleId).Select(m => m.TcfModuleName).FirstOrDefault(),
                                 tcfModuleId = wf.TcfModuleId
                             }).ToList();

                List<ViewTcfMaster> viewTcfMasterList = new List<ViewTcfMaster>();
                List<FirstAppCcLists> firstAppCcLists = new List<FirstAppCcLists>();
                List<SecondAppCcLists> secondAppCcLists = new List<SecondAppCcLists>();

                foreach (var item in check)
                {
                    ViewTcfMaster viewTcfMaster = new ViewTcfMaster();
                    viewTcfMaster.TcfApprovedMasterId = item.TcfApprovedMasterId;
                    if (item.FirstApproverToList != "" && item.FirstApproverToList != null)
                    {
                        viewTcfMaster.FirstApproverToList = item.FirstApproverToList;
                    }

                    if (item.FirstApproverCcList != null && item.FirstApproverCcList != "")
                    {
                        string[] ids = item.FirstApproverCcList.Split(',');

                        //foreach (var i in ids)
                        //{
                        //    FirstAppCcLists firstAppCcLists1 = new FirstAppCcLists();
                        //    firstAppCcLists1.FirstApproverCcList = i;
                        //    firstAppCcLists.Add(firstAppCcLists1);
                        //}
                        viewTcfMaster.FirstApproverCcList = ids;
                    }

                    if (item.SecondApproverToList != " " && item.SecondApproverToList != null)
                    {
                        viewTcfMaster.SecondApproverToList = item.SecondApproverToList;
                    }

                    if (item.SecondApproverCcList != null && item.SecondApproverCcList != "")
                    {
                        string[] ids1 = item.SecondApproverCcList.Split(',');

                        //foreach (var i in ids1)
                        //{
                        //    SecondAppCcLists secondAppCcLists1 = new SecondAppCcLists();
                        //    secondAppCcLists1.SecondApproverCcList = i;
                        //    secondAppCcLists.Add(secondAppCcLists1);
                        //}
                        viewTcfMaster.SecondApproverCcList = ids1;
                    }
                    viewTcfMaster.PlantId = Convert.ToInt32(item.plantId);
                    viewTcfMaster.ShopId = Convert.ToInt32(item.shopId);
                    viewTcfMaster.CellId = Convert.ToInt32(item.cellId);
                    viewTcfMaster.PlantName = item.plantName;
                    viewTcfMaster.ShopName = item.shopName;
                    viewTcfMaster.CellName = item.cellName;
                    viewTcfMaster.ModuleName = item.ModuleName;
                    viewTcfMaster.TcfModuleId = item.tcfModuleId;
                    viewTcfMasterList.Add(viewTcfMaster);
                }
                obj.isStatus = true;
                obj.response = viewTcfMasterList;
            }
            catch (Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isStatus = false;
            }
            return obj;
        }

        /// <summary>
        /// View Multiple Tcf Approved Master By Id
        /// </summary>
        /// <param name="tcfMasterId"></param>
        /// <returns></returns>
        public CommonResponse1 ViewMultipleTcfApprovedMasterById(int tcfMasterId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblTcfApprovedMaster
                             where wf.IsDeleted == 0 && wf.TcfApprovedMasterId == tcfMasterId
                             select new
                             {
                                 TcfApprovedMasterId = wf.TcfApprovedMasterId,
                                 FirstApproverToList = wf.FirstApproverToList,
                                 FirstApproverCcList = wf.FirstApproverCcList,
                                 SecondApproverToList = wf.SecondApproverToList,
                                 SecondApproverCcList = wf.SecondApproverCcList,
                                 plantName = db.Tblplant.Where(m => m.PlantId == wf.PlantId).Select(m => m.PlantName).FirstOrDefault(),
                                 plantId = wf.PlantId,
                                 shopName = db.Tblshop.Where(m => m.PlantId == wf.ShopId).Select(m => m.ShopName).FirstOrDefault(),
                                 shopId = wf.ShopId,
                                 cellName = db.Tblcell.Where(m => m.CellId == wf.CellId).Select(m => m.CellName).FirstOrDefault(),
                                 cellId = wf.CellId,
                                 ModuleName = db.TblTcfModule.Where(m => m.TcfModuleId == wf.TcfModuleId).Select(m => m.TcfModuleName).FirstOrDefault()
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
        /// Delete Multiple Tcf Approved Master
        /// </summary>
        /// <param name="tcfMasterId"></param>
        /// <returns></returns>
        public CommonResponse1 DeleteMultipleTcfApprovedMaster(int tcfMasterId)
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = db.TblTcfApprovedMaster.Where(m => m.TcfApprovedMasterId == tcfMasterId).FirstOrDefault();
                if (check != null)
                {
                    check.IsDeleted = 1;
                    check.ModifiedOn = DateTime.Now;
                    db.SaveChanges();
                    obj.isStatus = true;
                    obj.response = "Deleted Successfully";
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
        /// Get Module Names
        /// </summary>
        /// <returns></returns>
        public CommonResponse1 GetModules()
        {
            CommonResponse1 obj = new CommonResponse1();
            try
            {
                var check = (from wf in db.TblTcfModule
                             where wf.IsDeleted == 0 && wf.TcfModuleId != 6
                             select new
                             {
                                 TcfModuleId = wf.TcfModuleId,
                                 ModuleName = wf.TcfModuleName,
                                 TcfModuleDesc = wf.TcfModuleDesc
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

        //public CommonResponse AddUploadedApprovedMasterDetails(tcflist data)
        //{
        //    CommonResponse obj = new CommonResponse();
        //    try
        //    {
        //        foreach (var dataitem in data.approvedlist)
        //        {
        //            var plantid = db.Tblplant.Where(m => m.PlantName == dataitem.plantName).Select(m => m.PlantId).FirstOrDefault();
        //            if (plantid != 0)
        //            {
        //                var shopid = db.Tblshop.Where(m => m.ShopName == dataitem.shopName).Select(m => m.ShopId).FirstOrDefault();
        //                if (shopid != 0)
        //                {
        //                    if (dataitem.cellName != "")
        //                    {
        //                        var cellid = db.Tblcell.Where(m => m.CellName == dataitem.cellName).Select(m => m.CellId).FirstOrDefault();
        //                        if (cellid != 0)
        //                        {
        //                            var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
        //                            if (moduleid != 0)
        //                            {
        //                                var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.CellId == cellid && m.TcfModuleId == moduleid).FirstOrDefault();
        //                                if (actdet != null)
        //                                {
        //                                    actdet.PlantId = plantid;
        //                                    actdet.ShopId = shopid;
        //                                    actdet.CellId = cellid;
        //                                    actdet.TcfModuleId = moduleid;
        //                                    actdet.FirstApproverCcList = dataitem.firstCCList;
        //                                    actdet.FirstApproverToList = dataitem.firstToList;
        //                                    actdet.SecondApproverCcList = dataitem.secondCCList;
        //                                    actdet.SecondApproverToList = dataitem.secondToList;
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                                else
        //                                {
        //                                    TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
        //                                    prcobj.TcfModuleId = moduleid;
        //                                    prcobj.FirstApproverCcList = dataitem.firstCCList;
        //                                    prcobj.FirstApproverToList = dataitem.firstToList;
        //                                    prcobj.SecondApproverCcList = dataitem.secondCCList;
        //                                    prcobj.SecondApproverToList = dataitem.secondToList;
        //                                    prcobj.CellId = cellid;
        //                                    prcobj.ShopId = shopid;
        //                                    prcobj.PlantId = plantid;
        //                                    prcobj.CreatedBy = 1;
        //                                    prcobj.CreatedOn = DateTime.Now;
        //                                    prcobj.IsDeleted = 0;
        //                                    db.TblTcfApprovedMaster.Add(prcobj);
        //                                    db.SaveChanges();
        //                                    obj.isTure = true;
        //                                }
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        var moduleid = db.TblTcfModule.Where(m => m.TcfModuleName == dataitem.Modulename).Select(m => m.TcfModuleId).FirstOrDefault();
        //                        if (moduleid != 0)
        //                        {
        //                            var actdet = db.TblTcfApprovedMaster.Where(m => m.PlantId == plantid && m.ShopId == shopid && m.TcfModuleId == moduleid).FirstOrDefault();
        //                            if (actdet != null)
        //                            {
        //                                actdet.PlantId = plantid;
        //                                actdet.ShopId = shopid;
        //                                actdet.CellId = null;
        //                                actdet.TcfModuleId = moduleid;
        //                                actdet.FirstApproverCcList = dataitem.firstCCList;
        //                                actdet.FirstApproverToList = dataitem.firstToList;
        //                                actdet.SecondApproverCcList = dataitem.secondCCList;
        //                                actdet.SecondApproverToList = dataitem.secondToList;
        //                                db.SaveChanges();
        //                                obj.isTure = true;
        //                            }
        //                            else
        //                            {
        //                                TblTcfApprovedMaster prcobj = new TblTcfApprovedMaster();
        //                                prcobj.TcfModuleId = moduleid;
        //                                prcobj.FirstApproverCcList = dataitem.firstCCList;
        //                                prcobj.FirstApproverToList = dataitem.firstToList;
        //                                prcobj.SecondApproverCcList = dataitem.secondCCList;
        //                                prcobj.SecondApproverToList = dataitem.secondToList;
        //                                prcobj.CellId = null;
        //                                prcobj.ShopId = shopid;
        //                                prcobj.PlantId = plantid;
        //                                prcobj.CreatedBy = 1;
        //                                prcobj.CreatedOn = DateTime.Now;
        //                                prcobj.IsDeleted = 0;
        //                                db.TblTcfApprovedMaster.Add(prcobj);
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
