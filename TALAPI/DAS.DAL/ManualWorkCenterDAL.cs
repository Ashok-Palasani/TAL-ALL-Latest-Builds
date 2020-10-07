using DAS.DAL.Helpers;
using DAS.DBModels;
using DAS.Interface;
using Microsoft.Extensions.Options;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using static DAS.EntityModels.CommonEntity;
using static DAS.EntityModels.ManualWCEntity;
using DAS.EntityModels;

namespace DAS.DAL
{
    public class ManualWorkCenterDAL : IManualWorkCenter
    {

        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(ManualWorkCenterDAL));

        public ManualWorkCenterDAL(i_facility_talContext _db)
        {
            db = _db;
        }

        public CommonResponse GetPlants()
        {
            CommonResponse obj = new CommonResponse();

            var check = (from wf in db.Tblplant
                         where wf.IsDeleted == 0
                         select new
                         {
                             PlantId = wf.PlantId,
                             PlantName = wf.PlantName,
                             PlantDesc = wf.PlantDesc
                         }).ToList();
            if(check.Count > 0)
            {
                obj.isTure = true;
                obj.response = check;
            }
            else
            {
                obj.isTure = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        public CommonResponse GetShops(int plantId)
        {
            CommonResponse obj = new CommonResponse();
            var check = (from wf in db.Tblshop
                         where wf.IsDeleted == 0 && wf.PlantId == plantId
                         select new
                         {
                             wf.ShopId,
                             wf.ShopName,
                             wf.ShopDesc
                         }).ToList();
            if(check.Count > 0)
            {
                obj.isTure = true;
                obj.response = check;
            }
            else
            {
                obj.isTure = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        public CommonResponse GetCells(int shopId)
        {
            CommonResponse obj = new CommonResponse();
            var check = (from wf in db.Tblcell
                         where wf.IsDeleted == 0 && wf.ShopId == shopId
                         select new
                         {
                             wf.CellId,
                             wf.CellName,
                             wf.CellDesc
                         }).ToList();
            if(check.Count > 0)
            {
                obj.isTure = true;
                obj.response = check;
            }
            else
            {
                obj.isTure = false;
                obj.response = "No Items Found";
            }
            return obj;
        }

        public GeneralResponse AddStandardMachineDetails(AddMachineandUserDetails data)
        {
            GeneralResponse obj = new GeneralResponse();
            var check = db.Tblmachinedetails.Where(m => m.MachineInvNo == data.machineInvNo).FirstOrDefault();
            if(check == null)
            {
                Tblmachinedetails tblmachinedetails = new Tblmachinedetails();
                tblmachinedetails.MachineInvNo = data.machineInvNo;
                tblmachinedetails.MachineMake = data.machineMake;
                tblmachinedetails.MachineModel = data.machineModel;
                tblmachinedetails.ModelType = data.modelType;
                tblmachinedetails.PlantId = data.plantId;
                tblmachinedetails.ShopId = data.shopId;
                tblmachinedetails.CellId = data.cellId;
                tblmachinedetails.Ipaddress = data.ipaddress;
                tblmachinedetails.MachineType = "2";
                tblmachinedetails.IsPcb = data.isPcb;
                tblmachinedetails.ControllerType = data.controllerType;
                tblmachinedetails.MachineDispName = data.machineDispName;
                tblmachinedetails.InsertedOn = Convert.ToString(DateTime.Now);
                tblmachinedetails.InsertedBy = 1;
                tblmachinedetails.IsDeleted = 0;
                tblmachinedetails.IsParameters = 1;
                tblmachinedetails.IsNormalWc = 1;
                tblmachinedetails.IsSyncEnable = 0;
                tblmachinedetails.ShopNo = db.Tblshop.Where(m => m.ShopId == data.shopId).Select(m => m.ShopName).FirstOrDefault();
                db.Add(tblmachinedetails);
                db.SaveChanges();
                obj.isStatus = true;
                obj.response = "Standard WorkCenter Added Successfully";
                obj.machineId = tblmachinedetails.MachineId;
                obj.machineInvNo = tblmachinedetails.MachineInvNo;
            }

            else
            {
                obj.isStatus = false;
                obj.response = "This MachineInvNo is already Present";
            }
            return obj;
        }

        public CommonResponse GetManualWorkCenterCount(int no,string MachineInvNo)
        {
            List<MachineandUserDetails> manualWorkCenter = new List<MachineandUserDetails>();
            CommonResponse obj = new CommonResponse();
            try
            {
                int temp;
                var check = db.Tblmachinedetails.Where(m => m.MachineInvNo == MachineInvNo).FirstOrDefault();
                for (temp = no; temp > 0; temp--)
                {
                    string usernamePassword = MachineInvNo + "-0" + temp;
                    MachineandUserDetails manualWorkCenters = new MachineandUserDetails();
                    manualWorkCenters.machineId = check.MachineId;
                    manualWorkCenters.ipaddress = check.Ipaddress;
                    manualWorkCenters.controllerType = check.ControllerType;
                    manualWorkCenters.machineModel = check.MachineModel;
                    manualWorkCenters.machineMake = check.MachineMake;
                    manualWorkCenters.modelType = check.ModelType;
                    manualWorkCenters.machineDispName = check.MachineModel + "||" + MachineInvNo + "-0" + temp;
                    manualWorkCenters.isPcb = check.IsPcb;
                    manualWorkCenters.plantName = db.Tblplant.Where(m=>m.PlantId == check.PlantId).Select(m=>m.PlantName).FirstOrDefault();
                    manualWorkCenters.shopName = db.Tblshop.Where(m => m.ShopId == check.ShopId).Select(m=>m.ShopName).FirstOrDefault();
                    manualWorkCenters.cellName = db.Tblcell.Where(m=>m.CellId == check.CellId).Select(m=>m.CellName).FirstOrDefault();
                    manualWorkCenters.isSyncEnable = check.IsSyncEnable;
                    //manualWorkCenters.machineInvNo = MachineInvNo + "-0" + temp;
                    manualWorkCenters.machineInvNo = usernamePassword;
                    manualWorkCenters.plantId = check.PlantId;
                    manualWorkCenters.shopId = check.ShopId;
                    manualWorkCenters.cellId = check.CellId;

                    //string machineInvNo = check.MachineInvNo;
                    string machineInvNo = usernamePassword;
                    string[] macInvArry = machineInvNo.Split('-');
                    string mac = "";
                    int counter = 0;
                    foreach (var item in macInvArry)
                    {
                        if (counter == 0)
                        {
                            mac = item.ToLower();
                        }
                        else
                        {
                            mac = mac + item;
                        }
                        counter++;
                    }
                    manualWorkCenters.userName = mac;
                    manualWorkCenters.password = mac;
                    manualWorkCenter.Add(manualWorkCenters);  
                }
                obj.isTure = true;
                obj.response = manualWorkCenter.OrderBy(m=>m.machineInvNo);
            }
            catch(Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }

        public CommonResponse AddManualWorkCenterAndUserDetails(List<AddMachineandUserDetails> datas)
        {
            CommonResponse obj = new CommonResponse();
            try
            {
                foreach (var data in datas)
                {
                    var check = db.Tblmachinedetails.Where(m => m.MachineInvNo == data.machineInvNo).FirstOrDefault();
                    if (check == null)
                    {
                        Tblmachinedetails tblmachinedetails = new Tblmachinedetails();
                        tblmachinedetails.ManualWcid = data.manualWcid;
                        tblmachinedetails.MachineInvNo = data.machineInvNo;
                        tblmachinedetails.MachineMake = data.machineMake;
                        tblmachinedetails.MachineModel = data.machineModel;
                        tblmachinedetails.PlantId = data.plantId;
                        tblmachinedetails.ShopId = data.shopId;
                        tblmachinedetails.CellId = data.cellId;
                        tblmachinedetails.Ipaddress = data.ipaddress;
                        tblmachinedetails.ControllerType = data.controllerType;
                        tblmachinedetails.IsPcb = data.isPcb;
                        tblmachinedetails.MachineType = "2";
                        tblmachinedetails.MachineDispName = data.machineDispName;
                        tblmachinedetails.InsertedOn = Convert.ToString(DateTime.Now);
                        tblmachinedetails.InsertedBy = 1;
                        tblmachinedetails.IsDeleted = 0;
                        tblmachinedetails.IsParameters = 1;
                        tblmachinedetails.IsNormalWc = 1;
                        tblmachinedetails.IsSyncEnable = 0;
                        tblmachinedetails.ShopNo = db.Tblshop.Where(m => m.ShopId == data.shopId).Select(m => m.ShopName).FirstOrDefault();
                        db.Add(tblmachinedetails);
                        db.SaveChanges();
                        obj.isTure = true;
                        obj.response = "Manual WorkCenter Added Successfully";

                        //string machineInvNo = tblmachinedetails.MachineInvNo;
                        //string[] macInvArry = machineInvNo.Split('-');
                        //string mac = "";
                        //int counter = 0;
                        //foreach (var item in macInvArry)
                        //{
                        //    if (counter == 0)
                        //    {
                        //        mac = item.ToLower();
                        //    }
                        //    else {
                        //        mac = mac + item;
                        //    }
                        //    counter++;
                        //}

                        //string machineInvNoNew = mac;
                        //var dbCheck = db.Tblusers.Where(m => m.UserName == machineInvNoNew).FirstOrDefault();
                        //if (dbCheck == null)
                        //{
                        //    Tblusers tblusers = new Tblusers();
                        //    tblusers.UserName = machineInvNoNew;
                        //    tblusers.Password = machineInvNoNew;
                        //    tblusers.MachineId = tblmachinedetails.MachineId;
                        //    tblusers.CreatedBy = 1;
                        //    tblusers.DisplayName = "operator";
                        //    tblusers.CreatedOn = DateTime.Now;
                        //    tblusers.PrimaryRole = 3;
                        //    tblusers.SecondaryRole = 3;
                        //    tblusers.IsDeleted = 0;
                        //    db.Add(tblusers);
                        //    db.SaveChanges();
                        //    obj.isStatus = true;
                        //    obj.response = "User Created Successfully";
                        //}
                        int machineId = tblmachinedetails.MachineId;
                        var dbCheck = db.Tblusers.Where(m => m.MachineId == machineId).FirstOrDefault();
                        if(dbCheck==null)
                        {
                            Tblusers tblusers = new Tblusers();
                            tblusers.UserName = data.userName;
                            tblusers.Password = data.password;
                            tblusers.MachineId = machineId;
                            tblusers.CreatedBy = 1;
                            tblusers.DisplayName = "operator";
                            tblusers.CreatedOn = DateTime.Now;
                            tblusers.PrimaryRole = 3;
                            tblusers.SecondaryRole = 3;
                            tblusers.IsDeleted = 0;
                            db.Add(tblusers);
                            db.SaveChanges();
                            obj.isTure = true;
                            obj.response = "User Created Successfully";
                        }
                        else
                        {
                            dbCheck = db.Tblusers.Where(m => m.MachineId == machineId).FirstOrDefault();
                            dbCheck.UserName = data.userName;
                            dbCheck.Password = data.password;
                            db.SaveChanges();
                        }
                    }
                    else
                    {
                        obj.isTure = false;
                        obj.response = "Manual WorkCenter Already Exist";
                    }
                }
            }
            catch(Exception ex)
            {
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
                obj.isTure = false;
            }
            return obj;
        }
    }
}
