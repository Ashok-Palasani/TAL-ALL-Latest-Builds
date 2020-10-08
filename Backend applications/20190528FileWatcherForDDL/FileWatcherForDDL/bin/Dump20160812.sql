CREATE DATABASE  IF NOT EXISTS `mazakdaq` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `mazakdaq`;
-- MySQL dump 10.13  Distrib 5.7.9, for Win64 (x86_64)
--
-- Host: localhost    Database: mazakdaq
-- ------------------------------------------------------
-- Server version	5.6.17-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `tblddl`
--

DROP TABLE IF EXISTS `tblddl`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tblddl` (
  `DDLID` int(11) NOT NULL AUTO_INCREMENT,
  `WorkCenter` varchar(45) DEFAULT NULL,
  `WorkCenterDesc` varchar(100) DEFAULT NULL,
  `FunctionOfWC` varchar(200) DEFAULT NULL,
  `ProductionOrder` varchar(45) DEFAULT NULL,
  `OperationNo` varchar(45) DEFAULT NULL,
  `OperationShortDesc` varchar(100) DEFAULT NULL,
  `MaterialNumber` varchar(45) DEFAULT NULL,
  `MaterialDesc` varchar(150) DEFAULT NULL,
  `Project` varchar(45) DEFAULT NULL,
  `MADDate` date DEFAULT NULL,
  `MADDateInd` varchar(45) DEFAULT NULL,
  `PreOpnEndDate` date DEFAULT NULL,
  `DaysAgeing` int(10) DEFAULT NULL,
  `ScheStartDate` date DEFAULT NULL,
  `OrderType` varchar(45) DEFAULT NULL,
  `Qty` int(10) DEFAULT NULL,
  `DueDate` date DEFAULT NULL,
  `ECDDateInd` varchar(45) DEFAULT NULL,
  `FiagRushInd` varchar(45) DEFAULT NULL,
  `OperationsOnHold` varchar(45) DEFAULT NULL,
  `PreOpnWorkCenter` varchar(45) DEFAULT NULL,
  `Addition` varchar(10) DEFAULT NULL,
  `UniqueNo` varchar(100) DEFAULT NULL,
  `Duplicates` int(2) DEFAULT NULL,
  `ProdFAI` varchar(45) DEFAULT NULL,
  `IsCompleted` int(2) NOT NULL DEFAULT '0',
  `IsDeleted` int(2) NOT NULL DEFAULT '0',
  `DeliveredQty` int(10) NOT NULL DEFAULT '0',
  `ScrapQty` int(10) NOT NULL DEFAULT '0',
  PRIMARY KEY (`DDLID`),
  UNIQUE KEY `DDLID_UNIQUE` (`DDLID`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tblddl`
--

LOCK TABLES `tblddl` WRITE;
/*!40000 ALTER TABLE `tblddl` DISABLE KEYS */;
INSERT INTO `tblddl` VALUES (1,'MS-02-04','MAZAK_3AXIS-1_Travel',NULL,'14000000155','40','CNCMILLING-1','50006037','KH30408','ROLLSROYCE','2019-03-20','X','2016-03-28',103,'2028-03-20','ZROL',2,'2028-03-20',NULL,NULL,NULL,'RS-01','N','1400000015542457MS-02-04',1,NULL,0,0,0,0),(2,'MS-02-04','MAZAK_3AXIS-1_Travel',NULL,'14000000315','50','CNCMILLING-2','50006034','KH30404','ROLLSROYCE','2011-04-20','X','2016-06-22',29,'2020-04-20','ZROL',132,'2020-04-20',NULL,NULL,NULL,'MS-02-04','N','1400000031542543MS-02-04',1,NULL,0,0,0,0),(3,'MS-02-05','MAZAK_3AXIS-2_Travel',NULL,'14000000352','80','CNCMILLING-1','50006060','KH48476','GE','2017-10-20',NULL,'2016-07-20',5,'2005-08-20','ZDUM',1,'2005-08-20',NULL,NULL,NULL,'RS-01','N','600000077442571MS-02-04',1,NULL,0,0,0,0);
/*!40000 ALTER TABLE `tblddl` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tblnetworkdetailsforddl`
--

DROP TABLE IF EXISTS `tblnetworkdetailsforddl`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tblnetworkdetailsforddl` (
  `NPFDDLID` int(11) NOT NULL AUTO_INCREMENT,
  `Path` varchar(250) NOT NULL,
  `UserName` varchar(100) DEFAULT NULL,
  `Password` varchar(100) DEFAULT NULL,
  `DomainName` varchar(100) DEFAULT NULL,
  `IsDeleted` int(2) NOT NULL DEFAULT '0',
  `CreatedOn` datetime NOT NULL,
  `CreatedBy` int(2) NOT NULL,
  `ModifiedOn` datetime DEFAULT NULL,
  `ModifiedBy` int(2) DEFAULT NULL,
  PRIMARY KEY (`NPFDDLID`),
  UNIQUE KEY `NPFDDLID_UNIQUE` (`NPFDDLID`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tblnetworkdetailsforddl`
--

LOCK TABLES `tblnetworkdetailsforddl` WRITE;
/*!40000 ALTER TABLE `tblnetworkdetailsforddl` DISABLE KEYS */;
INSERT INTO `tblnetworkdetailsforddl` VALUES (1,'\\\\SRKS_TECH-2\\Users\\Tech-2\\J\\2016-03-16','SRKS_TECH-1','TECH-1','SRKS_TECH-1',1,'2016-08-12 10:10:00',1,NULL,NULL),(2,' \\\\tarafs\\UnitWorks\\DDL','unitworks','Unit@works','taslaero.com',0,'2016-08-12 10:10:00',1,NULL,NULL);
/*!40000 ALTER TABLE `tblnetworkdetailsforddl` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2016-08-12 21:06:10
