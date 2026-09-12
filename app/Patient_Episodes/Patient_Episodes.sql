CREATE TABLE IF NOT EXISTS `patient_episodes` (
  `RecordID` int NOT NULL AUTO_INCREMENT,
  `EpisodeUUID` char(36) NOT NULL,
  `PatientID` varchar(8) NOT NULL,
  `StartDate` date NOT NULL,
  `EndDate` date NOT NULL,
  `VisitTypeFlag` tinyint NOT NULL COMMENT '1:入院 2:外来',
  `DeleteFlag` tinyint NOT NULL DEFAULT '0',
  `InputDateTime` datetime NOT NULL,
  `InputStaffID` varchar(4) NOT NULL,
  `InputPC` varchar(10) NOT NULL,
  `UpdateDateTime` datetime NOT NULL,
  `UpdateStaffID` varchar(4) NOT NULL,
  `UpdatePC` varchar(10) NOT NULL,
  PRIMARY KEY (`RecordID`),
  UNIQUE KEY `EpisodeUUID` (`EpisodeUUID`),
  UNIQUE KEY `PatientID` (`PatientID`,`StartDate`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COMMENT='UUIDでFKを管理するCore Table' ;