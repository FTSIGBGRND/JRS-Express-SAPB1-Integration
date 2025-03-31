DELIMITER $$

USE `ftdbw_jrs`$$

DROP PROCEDURE IF EXISTS `FTSI_IMPORT_AR_INVOICE`$$

CREATE DEFINER=`root`@`localhost` PROCEDURE `FTSI_IMPORT_AR_INVOICE`(
	IN Id VARCHAR(36)
)
BEGIN

-- HEADER --
SELECT 	DocType, 
	U_RefNum,
	CardCode, 
	NumAtCard, 
	DATE_FORMAT(T1.DocDate, "%Y%m%d") 'DocDate',
	DATE_FORMAT(T1.DocDueDate, "%Y%m%d") 'DocDueDate',
	DATE_FORMAT(T1.TaxDate, "%Y%m%d") 'TaxDate',
	DiscPrcnt,
	Comments,
	U_JRSBranch,
	U_TransactionType,
	U_SalesType,
	U_ParentBP,
	U_AirwayBillNo,
	Id AS 'U_Id'

FROM ftoinv T1
WHERE T1.Id = Id;


-- LINES --
SELECT	LineNum,
	DESCRIPTION AS 'Dscription', 
	AccountCode AS 'AcctCode',
	Price AS 'PriceBefDi', 
	DiscPrcnt,
	VatGroup, 
	WTLiable AS 'WtLiable',
	PriceAfVAT, 
	OcrCode, 
	OcrCode2, 
	OcrCode3, 
	OcrCode4, 
	OcrCode5,
	U_xWTVendor,
	U_xTaxbleAmnt,
	U_xTaxAmnt,
	U_xSupplierName,
	U_xAddress,
	U_xTINnumber,
	U_xCardType,
	U_SalesType,
	U_RefNum

FROM ftinv1 T1
WHERE T1.HeaderId = Id;

-- WTax --
SELECT  U_RefNum,
	T1.WTCode,
	TaxbleAmnt,
	WTAmnt

FROM ftinv5 T1 
WHERE T1.HeaderId = Id AND 
	  IFNULL(WTCode,'') <> '';

	
END$$

DELIMITER ;