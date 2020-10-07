using DAS.DBModels;
using DAS.EntityModels;
using DAS.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static DAS.EntityModels.CommonResponseWithMachineName;

namespace DAS.DAL
{
    public class DALPlantShopCellData : IPlantShopCellData
    {
        i_facility_talContext db = new i_facility_talContext();
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(DALHMIDetails));

        public DALPlantShopCellData(i_facility_talContext _db)
        {
            db = _db;
        }

        public CommonResponse GetCellDetails(int shopId)
        {
            CommonResponse retData = new CommonResponse();
            List<CellEntity> listCellEntity = new List<CellEntity>();
            try
            {
                var shopData = db.Tblcell.Where(x => x.IsDeleted == 0 && x.ShopId == shopId).ToList();
                foreach (var row in shopData)
                {
                    CellEntity objCell = new CellEntity();
                    objCell.cellId = row.CellId;
                    objCell.cellName = row.CellName;
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

        public CommonResponse GetMachineDetails(int cellId)
        {
            CommonResponse retData = new CommonResponse();
            List<MachineEntity> listMachineEntity = new List<MachineEntity>();
            try
            {
                var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == cellId).Select(x => new { x.MachineId, x.MachineInvNo }).ToList();
                foreach (var row in machineData)
                {
                    MachineEntity objMachine = new MachineEntity();
                    objMachine.machineId = row.MachineId;
                    objMachine.machineName = row.MachineInvNo;
                    listMachineEntity.Add(objMachine);
                }
                if (machineData != null)
                {
                    retData.isTure = true;
                    retData.response = listMachineEntity;
                }
                else
                {
                    retData.isTure = true;
                    //.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetPlantDetails()
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                List<PlantEntity> listPlantEntity = new List<PlantEntity>();
                var plantData = db.Tblplant.Where(x => x.IsDeleted == 0).ToList();
                foreach (var row in plantData)
                {
                    PlantEntity obj = new PlantEntity();
                    obj.plantId = row.PlantId;
                    obj.plantName = row.PlantName;
                    listPlantEntity.Add(obj);
                }
                if (plantData != null)
                {
                    retData.isTure = true;
                    retData.response = listPlantEntity;
                }
                else
                {
                    retData.isTure = true;
                    //.response = ResourceResponse.NoItemsFound; ;
                }
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetShopDetails(int plantId)
        {
            CommonResponse retData = new CommonResponse();
            List<ShopEntity> listShopEntity = new List<ShopEntity>();
            try
            {
                var shopData = db.Tblshop.Where(x => x.IsDeleted == 0 && x.PlantId == plantId).ToList();
                foreach (var row in shopData)
                {
                    ShopEntity objShop = new ShopEntity();
                    objShop.shopId = row.ShopId;
                    objShop.shopName = row.ShopName;
                    listShopEntity.Add(objShop);
                }
                if (shopData != null)
                {
                    retData.isTure = true;
                    retData.response = listShopEntity;
                }
                else
                {
                    retData.isTure = true;
                    //retData.response = ResourceResponse.NoItemsFound;
                }
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetPlantShopCellMachineNames(PlantShopCellMachineId data)
        {
            CommonResponse retData = new CommonResponse();
            List<PlantShopCellMachineIdWithName> pscmdet = new List<PlantShopCellMachineIdWithName>();
            try
            {
                PlantShopCellMachineIdWithName pscmobj = new PlantShopCellMachineIdWithName();
                var plantData = db.Tblplant.Where(m => m.IsDeleted == 0 && m.PlantId == data.plantId).Select(m => m.PlantName).FirstOrDefault();
                pscmobj.plantId = data.plantId;
                pscmobj.plantName = plantData;

                var shopData = db.Tblshop.Where(m => m.IsDeleted == 0 && m.ShopId == data.shopId).Select(m => m.ShopName).FirstOrDefault();
                pscmobj.shopId = data.shopId;
                pscmobj.shopName = shopData;

                var cellData = db.Tblcell.Where(m => m.IsDeleted == 0 && m.CellId == data.cellId).Select(m => m.CellName).FirstOrDefault();
                pscmobj.cellId = data.cellId;
                pscmobj.cellName = cellData;

                var machineData = db.Tblmachinedetails.Where(m => m.IsDeleted == 0 && m.MachineId == data.machineId).Select(m => m.MachineInvNo).FirstOrDefault();
                pscmobj.machineId = data.machineId;
                pscmobj.machineName = machineData;
                retData.isTure = true;
                retData.response = pscmobj;
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetAllDetailsBasedOnPlantShopCell(PlantShopCellGet data)
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                PlantShopCellList plantShopCellList = new PlantShopCellList();
                List<PlantEntity> listPlantEntity = new List<PlantEntity>();
                var plantData = db.Tblplant.Where(x => x.IsDeleted == 0).ToList();
                foreach (var row in plantData)
                {
                    PlantEntity obj = new PlantEntity();
                    obj.plantId = row.PlantId;
                    obj.plantName = row.PlantName;
                    listPlantEntity.Add(obj);
                }
                if (plantData != null)
                {
                    plantShopCellList.plantList = listPlantEntity;
                    var shopData = db.Tblshop.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId).ToList();
                    List<ShopEntity> listShopEntity = new List<ShopEntity>();
                    foreach (var row in shopData)
                    {
                        ShopEntity objShop = new ShopEntity();
                        objShop.shopId = row.ShopId;
                        objShop.shopName = row.ShopName;
                        listShopEntity.Add(objShop);
                    }
                    if (shopData != null)
                    {
                        plantShopCellList.shopList = listShopEntity;
                        List<CellEntity> listCellEntity = new List<CellEntity>();
                        var cellData = db.Tblcell.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId).ToList();
                        foreach (var row in cellData)
                        {
                            CellEntity objCell = new CellEntity();
                            objCell.cellId = row.CellId;
                            objCell.cellName = row.CellName;
                            listCellEntity.Add(objCell);
                        }
                        if (cellData != null)
                        {
                            plantShopCellList.cellList = listCellEntity;
                        }
                    }

                }

                retData.isTure = true;
                retData.response = plantShopCellList;
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetAllDetailsBasedOnPlantShopCellMachine(PlantShopCellGet data)
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                PlantShopCellMachineList plantShopCellList = new PlantShopCellMachineList();
                List<PlantEntity> listPlantEntity = new List<PlantEntity>();
                var plantData = db.Tblplant.Where(x => x.IsDeleted == 0).ToList();
                foreach (var row in plantData)
                {
                    PlantEntity obj = new PlantEntity();
                    obj.plantId = row.PlantId;
                    obj.plantName = row.PlantName;
                    listPlantEntity.Add(obj);
                }
                if (plantData != null)
                {
                    plantShopCellList.plantList = listPlantEntity;
                    var shopData = db.Tblshop.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId).ToList();
                    List<ShopEntity> listShopEntity = new List<ShopEntity>();
                    foreach (var row in shopData)
                    {
                        ShopEntity objShop = new ShopEntity();
                        objShop.shopId = row.ShopId;
                        objShop.shopName = row.ShopName;
                        listShopEntity.Add(objShop);
                    }
                    if (shopData != null)
                    {
                        plantShopCellList.shopList = listShopEntity;
                        List<CellEntity> listCellEntity = new List<CellEntity>();
                        var cellData = db.Tblcell.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId).ToList();
                        foreach (var row in cellData)
                        {
                            CellEntity objCell = new CellEntity();
                            objCell.cellId = row.CellId;
                            objCell.cellName = row.CellName;
                            listCellEntity.Add(objCell);
                        }
                        if (cellData != null)
                        {
                            plantShopCellList.cellList = listCellEntity;
                            List<MachineEntity> listMachineEntity = new List<MachineEntity>();
                            var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId).ToList();
                            foreach (var row in machineData)
                            {
                                MachineEntity objMachine = new MachineEntity();
                                objMachine.machineId = row.MachineId;
                                objMachine.machineName = row.MachineInvNo;
                                listMachineEntity.Add(objMachine);
                            }
                            if (machineData != null)
                            {
                                plantShopCellList.machineList = listMachineEntity;
                            }
                        }
                    }

                }

                retData.isTure = true;
                retData.response = plantShopCellList;
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

        public CommonResponse GetAllDetailsBasedOnPlantShopCellProcess(PlantShopCellGet data)
        {
            CommonResponse retData = new CommonResponse();
            try
            {
                PlantShopCellMachineList plantShopCellList = new PlantShopCellMachineList();
                List<PlantEntity> listPlantEntity = new List<PlantEntity>();
                var plantData = db.Tblplant.Where(x => x.IsDeleted == 0).ToList();
                foreach (var row in plantData)
                {
                    PlantEntity obj = new PlantEntity();
                    obj.plantId = row.PlantId;
                    obj.plantName = row.PlantName;
                    listPlantEntity.Add(obj);
                }
                if (plantData != null)
                {
                    plantShopCellList.plantList = listPlantEntity;
                    var shopData = db.Tblshop.Where(x => x.IsDeleted == 0 && x.PlantId == data.plantId).ToList();
                    List<ShopEntity> listShopEntity = new List<ShopEntity>();
                    foreach (var row in shopData)
                    {
                        ShopEntity objShop = new ShopEntity();
                        objShop.shopId = row.ShopId;
                        objShop.shopName = row.ShopName;
                        listShopEntity.Add(objShop);
                    }
                    if (shopData != null)
                    {
                        plantShopCellList.shopList = listShopEntity;
                        List<CellEntity> listCellEntity = new List<CellEntity>();
                        var cellData = db.Tblcell.Where(x => x.IsDeleted == 0 && x.ShopId == data.shopId).ToList();
                        foreach (var row in cellData)
                        {
                            CellEntity objCell = new CellEntity();
                            objCell.cellId = row.CellId;
                            objCell.cellName = row.CellName;
                            listCellEntity.Add(objCell);
                        }
                        if (cellData != null)
                        {
                            plantShopCellList.cellList = listCellEntity;
                            List<MachineEntity> listMachineEntity = new List<MachineEntity>();
                            var machineData = db.Tblmachinedetails.Where(x => x.IsDeleted == 0 && x.CellId == data.cellId).ToList();
                            foreach (var row in machineData)
                            {
                                MachineEntity objMachine = new MachineEntity();
                                objMachine.machineId = row.MachineId;
                                objMachine.machineName = row.MachineInvNo;
                                listMachineEntity.Add(objMachine);
                            }
                            if (machineData != null)
                            {
                                plantShopCellList.machineList = listMachineEntity;
                                List<ProcessEntity> listProcessEntity = new List<ProcessEntity>();
                                var processData = db.Tblactivity.Where(x => x.Isdeleted == 0 && x.MachineId == data.machineId).ToList();
                                foreach (var row in processData)
                                {
                                    var processName = db.TblProcess.Where(m => m.ProcessId == row.ProcessId && m.Isdeleted == 0).FirstOrDefault();
                                    if (processName != null)
                                    {
                                        ProcessEntity objProcess = new ProcessEntity();
                                        objProcess.processId = (int)row.ProcessId;
                                        objProcess.processName = processName.ProcessName;
                                        listProcessEntity.Add(objProcess);
                                    }
                                }
                                if (processData != null)
                                {
                                    plantShopCellList.processList = listProcessEntity;
                                }
                            }

                        }
                    }

                }

                retData.isTure = true;
                retData.response = plantShopCellList;
            }
            catch (Exception ex)
            {
                retData.isTure = false;
                //retData.response = ResourceResponse.ExceptionMessage;
                log.Error(ex); if (ex.InnerException != null) { log.Error(ex.InnerException.ToString()); }
            }
            return retData;
        }

    }
}
