using SAPbobsCOM;
using System;
using System.Data;
using System.IO;

namespace FTSISAPB1iService
{
    public class ImportARInvoice
    {
        private static DateTime dteStart;
        private static string strTransType;
        private static string strSapDocNum;

        public static void _ImportARInvoice()
        {
            string strId = string.Empty;
            string strU_RefNum = string.Empty;
            string strXmlPath = string.Empty;
            string strPostDocEntry, strPostDocNum, strAcctCode, strFormatCode;
            string strMySQLTable;

            //integrationstatus : P - Pending , S - Success, E - Error
            //Posted :  N - No, Y - Yes

            try
            {
                dteStart = DateTime.Now;

                // Initialize Object Type.
                GlobalFunction.getObjType(13);
                strTransType = "Documents - " + GlobalVariable.strDocType;

                // Get All data for processing using Stored Procedure
                DataSet dsProcessData = SQLSettings.getDataFromMySQL(string.Format("CALL FTSI_IMPORT_GET_PROCESS_DATA({0})", GlobalVariable.intObjType));

                // Run process for each row
                foreach (DataRow oDataRow in dsProcessData.Tables[0].Rows)
                {
                    strId = oDataRow["Id"].ToString();
                    strMySQLTable = oDataRow["MySQLTable"].ToString();
                    strU_RefNum = oDataRow["U_RefNum"].ToString();

                    try
                    {
                        // Validation: Check if U_RefNum exists in OINV
                        if (GlobalFunction.checkRefNum(strU_RefNum, GlobalVariable.strTableHeader))
                        {
                            // Get Document Header and Line Details
                            DataSet dsBusinessObject = SQLSettings.getDataFromMySQL(string.Format("CALL FTSI_IMPORT_AR_INVOICE('{0}')", strId));

                            // Rename DataTables.
                            // NOTE: Make sure to rename DataTable because the names will be used as TAGS in XML file.
                            dsBusinessObject.Tables[0].TableName = "OINV";
                            dsBusinessObject.Tables[1].TableName = "INV1";
                            dsBusinessObject.Tables[2].TableName = "INV5";

                            for (int i = 0; i < dsBusinessObject.Tables["INV1"].Rows.Count; i++)
                            {
                                strFormatCode = dsBusinessObject.Tables[1].Rows[i]["AcctCode"].ToString();
                                strAcctCode = GlobalFunction.getCodebyId("OACT", strFormatCode, "AcctCode", "FormatCode");
                                dsBusinessObject.Tables["INV1"].Rows[i]["AcctCode"] = strAcctCode;
                            }

                            // Process XML File Creation
                            strXmlPath = GenerateFilePath(dsBusinessObject.Tables["OINV"].Rows[0]["U_RefNum"].ToString());
                            XMLGenerator.GenerateXMLFile(GlobalVariable.oObjectType, dsBusinessObject, strXmlPath);


                            // Start XML Import
                            StartCompanyTransaction();

                            if (ImportDocumentsXML.importTempXMLDocument(strXmlPath, strId))
                            {
                                // Get Posted DocEntry and DocNum
                                strPostDocEntry = GlobalVariable.oCompany.GetNewObjectKey().ToString();
                                strPostDocNum = GlobalFunction.getDocNum(GlobalVariable.intObjType, strPostDocEntry);

                                // Output to Integration Log
                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), Path.GetFileName(strXmlPath), strPostDocEntry, strPostDocNum, dteStart, "S", GlobalVariable.intObjType.ToString(), string.Format("Successfully Posted {0}", strTransType));

                                // Update Staging DB
                                SQLSettings.executeQuery(string.Format("UPDATE {0} SET IntegrationStatus = 'S', DocNum = {1}, DocEntry = {2}, IntegrationMessage = \"Successfully Posted\" WHERE Id = '{3}'", strMySQLTable, strPostDocNum, strPostDocEntry, strId));

                                EndCompanyTransaction(BoWfTransOpt.wf_Commit);
                            }
                            else
                            {
                                // Output to Integration Log
                                // -- to be fill, docNum and docEntry
                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), Path.GetFileName(strXmlPath), "", "", dteStart, "E", "-" + GlobalVariable.intObjType.ToString(), string.Format("Error Posting SAP Business Object: {0}", GlobalVariable.strErrMsg.Replace("\\", "").Replace("\"", "'")));

                                // Update Staging DB
                                SQLSettings.executeQuery(string.Format("UPDATE {0} SET IntegrationStatus = 'E', IntegrationMessage = \"{1}\" WHERE Id = '{2}'", strMySQLTable, GlobalVariable.strErrMsg.Replace("\\", "").Replace("\"", "'"), strId));

                                EndCompanyTransaction(BoWfTransOpt.wf_RollBack);
                            }
                        }
                        else
                        {
                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), Path.GetFileName(strXmlPath), "", strU_RefNum, dteStart, "E", "", $"Validation failed: U_RefNum '{strU_RefNum}' already exists.");

                            // Update Staging DB
                            SQLSettings.executeQuery(string.Format("UPDATE {0} SET IntegrationStatus = 'E', IntegrationMessage = \"U_RefNum already exist\" WHERE Id = '{1}'", strMySQLTable, strId));
                        }
                    }
                    catch (Exception ex)
                    {
                        GlobalVariable.intErrNum = -111;
                        GlobalVariable.strErrMsg = string.Format("Error Processing Import. {0}", ex.Message.ToString());

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), Path.GetFileName(strXmlPath), "", strU_RefNum, dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        // Update Staging DB
                        SQLSettings.executeQuery(string.Format("UPDATE {0} SET IntegrationStatus = 'E', IntegrationMessage = \"{1}\" WHERE Id = '{2}'", strMySQLTable, GlobalVariable.strErrMsg.Replace("\\", "").Replace("\"", "'"), strId));

                        GC.Collect();

                        EndCompanyTransaction(BoWfTransOpt.wf_RollBack);
                    }
                   
                }

            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = string.Format("Error Processing Import. {0}", ex.Message.ToString());

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), Path.GetFileName(strXmlPath), "", strU_RefNum, dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                EndCompanyTransaction(BoWfTransOpt.wf_RollBack);
            }
        }
        private static void StartCompanyTransaction()
        {
            if (!(GlobalVariable.oCompany.InTransaction))
                GlobalVariable.oCompany.StartTransaction();
        }

        private static void EndCompanyTransaction(BoWfTransOpt transOpt)
        {
            if (GlobalVariable.oCompany.InTransaction)
                GlobalVariable.oCompany.EndTransaction(transOpt);
        }

        private static string GenerateFilePath(string strRefNum)
        {
            return GlobalVariable.strTempPath + string.Format("{0}_DOC_{1}_{2}_{3}_1.xml", GlobalVariable.strCompany, GlobalVariable.strTableHeader, GlobalVariable.intObjType, strRefNum, DateTime.Today.ToString("MMddyyyy"));
        }
    }
}
