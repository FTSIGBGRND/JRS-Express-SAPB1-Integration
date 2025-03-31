using System;
using System.IO;

namespace FTSISAPB1iService
{
    class SystemInitialization
    {
        public static bool initTables()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            if (SystemFunction.createUDT("FTISL", "FT Integration Service Log", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (SystemFunction.createUDT("FTISS", "FT Integration Service SetUp", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            return true;
        }
        public static bool initFields()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            #region "FRAMEWORK UDF"

            /************************** MARKETING DOCUMENTS ****************************************************************/

            if (SystemFunction.isUDFexists("OINV", "isExtract") == false)
                if (SystemFunction.createUDF("OINV", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "FileName") == false)
                if (SystemFunction.createUDF("OINV", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "RefNum") == false)
                if (SystemFunction.createUDF("OINV", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "RefNum") == false)
                if (SystemFunction.createUDF("INV1", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseLine") == false)
                if (SystemFunction.createUDF("INV1", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseRef") == false)
                if (SystemFunction.createUDF("INV1", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseType") == false)
                if (SystemFunction.createUDF("INV1", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "RefNum") == false)
                if (SystemFunction.createUDF("INV3", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseLine") == false)
                if (SystemFunction.createUDF("INV3", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseRef") == false)
                if (SystemFunction.createUDF("INV3", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseType") == false)
                if (SystemFunction.createUDF("INV3", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV5", "RefNum") == false)
                if (SystemFunction.createUDF("INV5", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            /************************** ITEM MASTER DATA ***************************************************************/

            if (SystemFunction.isUDFexists("OITM", "isExtract") == false)
                if (SystemFunction.createUDF("OITM", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;


            if (SystemFunction.isUDFexists("OITM", "RefCode") == false)
                if (SystemFunction.createUDF("OITM", "RefCode", "Reference Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            /************************** BUSINESS PARTNER DATA **********************************************************/

            if (SystemFunction.isUDFexists("OCRD", "isExtract") == false)
                if (SystemFunction.createUDF("OCRD", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRD", "RefCode") == false)
                if (SystemFunction.createUDF("OCRD", "RefCode", "Reference Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;


            /************************** INCOMING PAYMENT **********************************************************/

            if (SystemFunction.isUDFexists("ORCT", "isExtract") == false)
                if (SystemFunction.createUDF("ORCT", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "RefNum") == false)
                if (SystemFunction.createUDF("ORCT", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "FileName") == false)
                if (SystemFunction.createUDF("ORCT", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "", "") == false)
                    return false;

            /************************** ADMINISTRATION ****************************************************************/

            if (SystemFunction.isUDFexists("OUSR", "IntMsg") == false)
                if (SystemFunction.createUDF("OUSR", "IntMsg", "Integration Message", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            #endregion

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            #region AR INVOICE

            if (SystemFunction.isUDFexists("OINV", "Id") == false)
                if (SystemFunction.createUDF("OINV", "Id", "Id", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "JRSBranch") == false)
                if (SystemFunction.createUDF("OINV", "JRSBranch", "JRSBranch", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "TransactionType") == false)
                if (SystemFunction.createUDF("OINV", "TransactionType", "TransactionType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "SalesType") == false)
                if (SystemFunction.createUDF("OINV", "SalesType", "SalesType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "ParentBP") == false)
                if (SystemFunction.createUDF("OINV", "ParentBP", "ParentBP", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "AirwayBillNo") == false)
                if (SystemFunction.createUDF("OINV", "AirwayBillNo", "AirwayBillNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "SalesType") == false)
                if (SystemFunction.createUDF("INV1", "SalesType", "SalesType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;


            #endregion

            #region AR CREDIT MEMO

            if (SystemFunction.isUDFexists("ORIN", "Id") == false)
                if (SystemFunction.createUDF("ORIN", "Id", "Id", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "RefNum") == false)
                if (SystemFunction.createUDF("ORIN", "RefNum", "RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "JRSBranch") == false)
                if (SystemFunction.createUDF("ORIN", "JRSBranch", "JRSBranch", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "TransactionType") == false)
                if (SystemFunction.createUDF("ORIN", "TransactionType", "TransactionType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "SalesType") == false)
                if (SystemFunction.createUDF("ORIN", "SalesType", "SalesType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "ParentBP") == false)
                if (SystemFunction.createUDF("ORIN", "ParentBP", "ParentBP", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORIN", "AirwayBillNo") == false)
                if (SystemFunction.createUDF("ORIN", "AirwayBillNo", "AirwayBillNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;


            if (SystemFunction.isUDFexists("RIN1", "SalesType") == false)
                if (SystemFunction.createUDF("RIN1", "SalesType", "SalesType", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("RIN5", "RefNum") == false)
                if (SystemFunction.createUDF("RIN5", "RefNum", "RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            #endregion

            #region INCOMING PAYMENT

            if (SystemFunction.isUDFexists("ORCT", "Id") == false)
                if (SystemFunction.createUDF("ORCT", "Id", "Id", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("ORCT", "RefNum") == false)
                if (SystemFunction.createUDF("ORCT", "RefNum", "RefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("RCT2", "InvRefNum") == false)
                if (SystemFunction.createUDF("RCT2", "InvRefNum", "InvRefNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 150, "", "", "") == false)
                    return false;

            #endregion


            return true;
        }

        public static bool initUDO()
        {
            return true;
        }
        public static bool initFolders()
        {
            try
            {
                string strDate = DateTime.Today.ToString("MMddyyyy") + @"\";

                string strExp = @"Export\" + strDate;
                string strImp = @"Import\" + strDate;

                GlobalVariable.strErrLogPath = GlobalVariable.strFilePath + @"\Error Log";
                if (!Directory.Exists(GlobalVariable.strErrLogPath))
                    Directory.CreateDirectory(GlobalVariable.strErrLogPath);

                GlobalVariable.strSQLScriptPath = GlobalVariable.strFilePath + @"\SQL Scripts\";
                if (!Directory.Exists(GlobalVariable.strSQLScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSQLScriptPath);

                GlobalVariable.strSAPScriptPath = GlobalVariable.strFilePath + @"\SAP Scripts\";
                if (!Directory.Exists(GlobalVariable.strSAPScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSAPScriptPath);

                GlobalVariable.strExpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strExpSucPath);

                GlobalVariable.strExpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strExpErrPath);

                GlobalVariable.strImpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strImpSucPath);

                GlobalVariable.strImpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strImpErrPath);

                GlobalVariable.strImpPath = GlobalVariable.strFilePath + @"\Import Files\";
                if (!Directory.Exists(GlobalVariable.strImpPath))
                    Directory.CreateDirectory(GlobalVariable.strImpPath);

                GlobalVariable.strExpPath = GlobalVariable.strFilePath + @"\Export Files\";
                if (!Directory.Exists(GlobalVariable.strExpPath))
                    Directory.CreateDirectory(GlobalVariable.strExpPath);

                GlobalVariable.strConPath = GlobalVariable.strFilePath + @"\Connection Path\";
                if (!Directory.Exists(GlobalVariable.strConPath))
                    Directory.CreateDirectory(GlobalVariable.strConPath);

                GlobalVariable.strTempPath = GlobalVariable.strFilePath + @"\Temp Files\";
                if (!Directory.Exists(GlobalVariable.strTempPath))
                    Directory.CreateDirectory(GlobalVariable.strTempPath);

                GlobalVariable.strAttImpPath = GlobalVariable.strFilePath + @"\Attachment\" + strImp;
                if (!Directory.Exists(GlobalVariable.strAttImpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttImpPath);

                GlobalVariable.strAttExpPath = GlobalVariable.strFilePath + @"\Attachment\" + strExp;
                if (!Directory.Exists(GlobalVariable.strAttExpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttExpPath);

                GlobalVariable.strArcExpPath = GlobalVariable.strFilePath + @"\Archive Files\Export\";
                if (!Directory.Exists(GlobalVariable.strArcExpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcExpPath);

                GlobalVariable.strArcImpPath = GlobalVariable.strFilePath + @"\Archive Files\Import\";
                if (!Directory.Exists(GlobalVariable.strArcImpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcImpPath);

                GlobalVariable.strCRPath = GlobalVariable.strFilePath + @"\Crystal Report\";
                if (!Directory.Exists(GlobalVariable.strCRPath))
                    Directory.CreateDirectory(GlobalVariable.strCRPath);


                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(string.Format("Error initializing program directory. {0}", ex.Message.ToString()));
                return false;
            }
        }
        public static bool initStoreProcedure()
        {
            //if (!(SystemFunction.initStoredProcedures(GlobalVariable.strSAPScriptPath)))
            //    return false;

            return true;
        }
    }
}
