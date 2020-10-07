using System;
using System.Collections.Generic;
using System.Text;

namespace DAS.EntityModels
{
    public class EntityModel
    {
        public bool isTrue { get; set; }
        public dynamic response { get; set; }
    }

    public class EntityModelWithLossCount
    {
        public bool isTrue { get; set; }
        public dynamic response { get; set; }
        public int count { get; set; }
        public dynamic errorMsg { get; set; }
    }

    public class CommonResponseWithTempId
    {
        public bool isStatus { get; set; }
        public dynamic response { get; set; }
        public int tempModeId { get; set; }
    }
   

    public class Lossdet
    {
        public int machineID { get; set; }
        public string machineNmae { get; set; }
        public int NCID { get; set; }
        //public string MessageTypeID { get; set; }
        public string startTime { get; set; }
        public string endTime { get; set; }
        public int Reasonlevel1Id { get; set; }
        public string Reasonlevel1Name { get; set; }
        public int Reasonlevel2Id { get; set; }
        public string Reasonlevel2Name { get; set; }
        public int Reasonlevel3Id { get; set; }
        public string Reasonlevel3Name { get; set; }
        //public double duration { get; set; }
    }
    //For index data
    public class LossdetData
    {

        //public string startTime { get; set; }
        //public string endTime { get; set; }
        //public int Reasonlevel1Id { get; set; }
        //public string Reasonlevel1Name { get; set; }
        //public int Reasonlevel2Id { get; set; }
        //public string Reasonlevel2Name { get; set; }
        //public int Reasonlevel3Id { get; set; }
        //public string Reasonlevel3Name { get; set; }
        public int machineID { get; set; }
        public string machineNmae { get; set; }
        public int NCID { get; set; }
        public string plantName { get; set; }
        public string shopName { get; set; }
        public string cellName { get; set; }
        public string CorrectedDate { get; set; }
        public string firstApproval { get; set; }
        public string secondApproval { get; set; }
    }


    public class lossmodel
    {

        public int plantId { get; set; }
        public int shopId { get; set; }
        public int cellId { get; set; }
        public int machineId { get; set; }
        public string FromDate { get; set; }
        //public int skipValue { get; set; }
        //public int takeValue { get; set; }
        //public string ToDate { get; set; }
    }
}
