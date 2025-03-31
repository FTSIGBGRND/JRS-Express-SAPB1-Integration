DELIMITER $$

USE `ftdbw_jrs`$$

DROP PROCEDURE IF EXISTS `FTSI_IMPORT_GET_PROCESS_DATA`$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `FTSI_IMPORT_GET_PROCESS_DATA`(
  IN ObjType INT
)
BEGIN

-- AR INVOICE --
IF ObjType = 13 
THEN 
	SELECT 'ftoinv' AS MySQLTable,
	        Id,
	        U_RefNum
	FROM ftoinv
	WHERE IFNULL(IntegrationStatus, 'P') = 'P';
END IF;

-- AR CREDIT MEMO --
IF ObjType = 14
THEN 
	SELECT 'ftorin' AS MySQLTable,
		Id,
		U_RefNum
	FROM ftorin 
	WHERE IFNULL(IntegrationStatus, 'P') = 'P';
END IF;

-- INCOMING PAYMENT --
IF ObjType = 24 
THEN 
	SELECT 'ftorct' AS MySQLTable,
		T1.Id,
		T1.U_RefNum
	FROM ftorct T1
	WHERE IFNULL(CardCode, '') <> '' 
	AND IFNULL(IntegrationStatus, 'P') = 'P';
END IF;
END$$

DELIMITER ;