DELIMITER $$

USE `ftdbw_jrs`$$

DROP PROCEDURE IF EXISTS `FTSI_IMPORT_INCOMING_PAYMENT`$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `FTSI_IMPORT_INCOMING_PAYMENT`(
	IN Id VARCHAR(36)
)
BEGIN

-- HEADER --
SELECT  CardCode,
	-- CashAccnt as 'CashAcct',
	DATE_FORMAT(T1.DocDate, "%Y%m%d") 'DocDate',
	-- CashSum, 
	TrsfrAcct,
	DATE_FORMAT(T1.TrsfrDate, "%Y%m%d") 'TrsfrDate',
	TrsfrRef, 
	TrsfrSum,
	U_RefNum, 
	Id AS 'U_Id'

FROM ftorct T1
WHERE T1.Id = Id;

-- CHECKS TABLE --

SELECT	BankCode,
	CheckNum,
	CHECKSUM AS 'CheckSum',
	DATE_FORMAT(T1.DueDate, "%Y%m%d") 'DueDate'

FROM ftrct1 T1
WHERE T1.Id = Id AND 
	  IFNULL(CheckNum, 0) <> '';

-- INVOICE TABLE --
SELECT	InvType,
	T1.U_RefNum AS 'U_InvRefNum',
	T2.DocEntry AS 'DocEntry'
	
FROM ftrct2 T1
LEFT JOIN ftoinv T2
	ON T1.U_RefNum = T2.U_RefNum
WHERE T1.Id = Id;

-- CREDIT VOUCHERS TABLE --
SELECT	LineID,
	CreditCard,
	CreditAcct,
	CrCardNum,
	DATE_FORMAT(T1.CardValid, "%Y%m%d") 'CardValid',
	VoucherNum,
	CreditSum,
	CrTypeCode

FROM ftrct3 T1
WHERE T1.Id = Id AND 
	  IFNULL(CreditCard, 0) <> '';


END$$

DELIMITER ;