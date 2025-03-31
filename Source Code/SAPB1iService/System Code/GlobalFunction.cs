﻿using PgpCore;
using SAPbobsCOM;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Xml;

namespace FTSISAPB1iService
{
    class GlobalFunction
    {
        public static void getObjType(int ObjType)
        {
            switch (ObjType)
            {
                case 2:
                    GlobalVariable.strDocType = "BP Master Data";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oBusinessPartners;
                    GlobalVariable.intObjType = 2;
                    GlobalVariable.strTableHeader = "OCRD";
                    GlobalVariable.strTableLine1 = "CRD1";
                    GlobalVariable.strTableLine3 = "OCPR";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 4:
                    GlobalVariable.strDocType = "Item Master Data";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oItems;
                    GlobalVariable.intObjType = 4;
                    GlobalVariable.strTableHeader = "OITM";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 13:
                    GlobalVariable.strDocType = "AR Invoice";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInvoices;
                    GlobalVariable.intObjType = 13;
                    GlobalVariable.strTableHeader = "OINV";
                    GlobalVariable.strTableLine1 = "INV1";
                    GlobalVariable.strTableLine3 = "INV3";
                    GlobalVariable.strTableLine5 = "INV5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 14:
                    GlobalVariable.strDocType = "AR Credit Memo";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes;
                    GlobalVariable.intObjType = 14;
                    GlobalVariable.strTableHeader = "ORIN";
                    GlobalVariable.strTableLine1 = "RIN1";
                    GlobalVariable.strTableLine3 = "RIN3";
                    GlobalVariable.strTableLine5 = "RIN5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 15:
                    GlobalVariable.strDocType = "Delivery";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                    GlobalVariable.intObjType = 15;
                    GlobalVariable.strTableHeader = "ODLN";
                    GlobalVariable.strTableLine1 = "DLN1";
                    GlobalVariable.strTableLine3 = "DLN3";
                    GlobalVariable.strTableLine5 = "DLN5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 16:
                    GlobalVariable.strDocType = "Sales Return";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oReturns;
                    GlobalVariable.intObjType = 16;
                    GlobalVariable.strTableHeader = "ORDN";
                    GlobalVariable.strTableLine1 = "RDN1";
                    GlobalVariable.strTableLine3 = "RDN3";
                    GlobalVariable.strTableLine5 = "RDN5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 17:
                    GlobalVariable.strDocType = "Sales Order";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oOrders;
                    GlobalVariable.intObjType = 17;
                    GlobalVariable.strTableHeader = "ORDR";
                    GlobalVariable.strTableLine1 = "RDR1";
                    GlobalVariable.strTableLine3 = "RDR3";
                    GlobalVariable.strTableLine5 = "RDR5";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 18:
                    GlobalVariable.strDocType = "AP Invoice";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                    GlobalVariable.intObjType = 18;
                    GlobalVariable.strTableHeader = "OPCH";
                    GlobalVariable.strTableLine1 = "PCH1";
                    GlobalVariable.strTableLine3 = "PCH3";
                    GlobalVariable.strTableLine5 = "PCH5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 19:
                    GlobalVariable.strDocType = "AP Credit Memo";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                    GlobalVariable.intObjType = 19;
                    GlobalVariable.strTableHeader = "ORPC";
                    GlobalVariable.strTableLine1 = "RPC1";
                    GlobalVariable.strTableLine3 = "RPC3";
                    GlobalVariable.strTableLine5 = "RPC5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 20:
                    GlobalVariable.strDocType = "Goods Receipt PO";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    GlobalVariable.intObjType = 20;
                    GlobalVariable.strTableHeader = "OPDN";
                    GlobalVariable.strTableLine1 = "PDN1";
                    GlobalVariable.strTableLine3 = "PDN3";
                    GlobalVariable.strTableLine5 = "PDN5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 21:
                    GlobalVariable.strDocType = "Goods Return";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    GlobalVariable.intObjType = 21;
                    GlobalVariable.strTableHeader = "ORPD";
                    GlobalVariable.strTableLine1 = "RPD1";
                    GlobalVariable.strTableLine3 = "RPD3";
                    GlobalVariable.strTableLine5 = "RPD5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 22:
                    GlobalVariable.strDocType = "Purchase Order";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                    GlobalVariable.intObjType = 22;
                    GlobalVariable.strTableHeader = "OPOR";
                    GlobalVariable.strTableLine1 = "POR1";
                    GlobalVariable.strTableLine3 = "POR3";
                    GlobalVariable.strTableLine5 = "POR5";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 23:
                    GlobalVariable.strDocType = "Sales Quotations";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseQuotations;
                    GlobalVariable.strTableHeader = "OQUT";
                    GlobalVariable.strTableLine1 = "QUT1";
                    GlobalVariable.strTableLine3 = "QUT3";
                    GlobalVariable.strTableLine5 = "QUT5";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 24:
                    GlobalVariable.strDocType = "Incoming Payment";
                    GlobalVariable.intObjType = 24;
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oIncomingPayments;
                    GlobalVariable.strTableHeader = "ORCT";
                    GlobalVariable.strTableLine1 = "RCT1";
                    GlobalVariable.strTableLine2 = "RCT2";
                    GlobalVariable.strTableLine3 = "RCT3";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 30:
                    GlobalVariable.strDocType = "Journal Entry";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oJournalEntries;
                    GlobalVariable.intObjType = 30;
                    GlobalVariable.strTableHeader = "OJDT";
                    GlobalVariable.strTableLine1 = "JDT1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 46:
                    GlobalVariable.strDocType = "Outgoing Payment";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oVendorPayments;
                    GlobalVariable.strTableHeader = "OVPM";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 59:
                    GlobalVariable.strDocType = "Goods Receipt";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
                    GlobalVariable.intObjType = 59;
                    GlobalVariable.strTableHeader = "OIGN";
                    GlobalVariable.strTableLine1 = "IGN1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 60:
                    GlobalVariable.strDocType = "Goods Issue";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;
                    GlobalVariable.intObjType = 60;
                    GlobalVariable.strTableHeader = "OIGE";
                    GlobalVariable.strTableLine1 = "IGE1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 67:
                    GlobalVariable.strDocType = "Stock Transfer";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    GlobalVariable.strTableHeader = "OWTR";
                    GlobalVariable.strTableLine1 = "WTR1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 66:
                    GlobalVariable.strDocType = "Bill of Materials";
                    GlobalVariable.intObjType = 66;
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oProductTrees;
                    GlobalVariable.strTableHeader = "OITT";
                    GlobalVariable.strTableLine1 = "ITT1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 112:
                    GlobalVariable.strDocType = "Draft";
                    GlobalVariable.intObjType = 112;
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oDrafts;
                    GlobalVariable.strTableHeader = "ODRF";
                    GlobalVariable.strTableLine1 = "DRF1";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 203:
                    GlobalVariable.strDocType = "AR DownPayment";
                    GlobalVariable.intObjType = 203;
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oDownPayments;
                    GlobalVariable.strTableHeader = "ODPI";
                    GlobalVariable.strTableLine1 = "DPI1";
                    GlobalVariable.strTableLine1 = "DPI5";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                case 204:
                    GlobalVariable.strDocType = "AP DownPayment";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments;
                    GlobalVariable.strTableHeader = "ODPO";
                    GlobalVariable.strTableLine1 = "DPO1";
                    GlobalVariable.blAlwUpdte = false;
                    break;

               

                case 1250000001:
                    GlobalVariable.strDocType = "Stock Transfer Request";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
                    GlobalVariable.strTableHeader = "OWTQ";
                    GlobalVariable.strTableLine1 = "WTQ1";
                    GlobalVariable.blAlwUpdte = true;
                    break;


                case 310:
                    GlobalVariable.strDocType = "Fixed Asset Retirement";
                    GlobalVariable.intObjType = 310;
                    //GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes
                    GlobalVariable.strTableHeader = "ORTI";
                    GlobalVariable.strTableLine1 = "RTI1";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 1470000049:
                    GlobalVariable.strDocType = "Fixed Asset Capitalization";
                    GlobalVariable.intObjType = 1470000049;
                    //GlobalVariable.oObjectType =
                    GlobalVariable.strTableHeader = "OACQ";
                    GlobalVariable.strTableLine1 = "ACQ1";
                    GlobalVariable.blAlwUpdte = true;
                    break;



                case 1470000113:
                    GlobalVariable.strDocType = "Purchase Request";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseRequest;
                    GlobalVariable.strTableHeader = "OPRQ";
                    GlobalVariable.strTableLine1 = "PRQ1";
                    GlobalVariable.blAlwUpdte = true;
                    break;

                case 28:
                    GlobalVariable.strDocType = "Journal Voucher";
                    GlobalVariable.oObjectType = SAPbobsCOM.BoObjectTypes.oJournalVouchers;
                    GlobalVariable.intObjType = 28;
                    GlobalVariable.strTableHeader = "OBTD";
                    GlobalVariable.blAlwUpdte = false;
                    break;

                default:
                    GlobalVariable.strDocType = "";
                    GlobalVariable.oObjectType = 0;
                    GlobalVariable.strTableHeader = "";
                    GlobalVariable.blAlwUpdte = false;
                    break;

            }
        }
        public static void getBaseType(int ObjType)
        {
            switch (ObjType)
            {
                case 13:
                    GlobalVariable.strBDocType = "AR Invoice";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInvoices;
                    GlobalVariable.intBObjType = 13;
                    GlobalVariable.strBTableHeader = "OINV";
                    GlobalVariable.strBTableLine1 = "INV1";
                    GlobalVariable.strBTableLine3 = "INV3";
                    GlobalVariable.strBTableLine5 = "INV5";
                    break;

                case 14:
                    GlobalVariable.strBDocType = "AR Credit Memo";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oCreditNotes;
                    GlobalVariable.intBObjType = 14;
                    GlobalVariable.strBTableHeader = "ORIN";
                    GlobalVariable.strBTableLine1 = "RIN1";
                    GlobalVariable.strBTableLine3 = "RIN3";
                    GlobalVariable.strBTableLine5 = "RIN5";
                    break;

                case 15:
                    GlobalVariable.strBDocType = "Delivery";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                    GlobalVariable.intBObjType = 15;
                    GlobalVariable.strBTableHeader = "ODLN";
                    GlobalVariable.strBTableLine1 = "DLN1";
                    GlobalVariable.strBTableLine3 = "DLN3";
                    GlobalVariable.strBTableLine5 = "DLN5";
                    break;

                case 16:
                    GlobalVariable.strBDocType = "Sales Return";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oReturns;
                    GlobalVariable.intBObjType = 16;
                    GlobalVariable.strBTableHeader = "ORDN";
                    GlobalVariable.strBTableLine1 = "RDN1";
                    GlobalVariable.strBTableLine3 = "RDN3";
                    GlobalVariable.strBTableLine5 = "RDN5";

                    break;

                case 17:
                    GlobalVariable.strBDocType = "Sales Order";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oOrders;
                    GlobalVariable.intBObjType = 17;
                    GlobalVariable.strBTableHeader = "ORDR";
                    GlobalVariable.strBTableLine1 = "RDR1";
                    GlobalVariable.strBTableLine3 = "RDR3";
                    GlobalVariable.strBTableLine5 = "RDR5";
                    break;

                case 18:
                    GlobalVariable.strBDocType = "AP Invoice";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                    GlobalVariable.intBObjType = 18;
                    GlobalVariable.strBTableHeader = "OPCH";
                    GlobalVariable.strBTableLine1 = "PCH1";
                    GlobalVariable.strBTableLine3 = "PCH3";
                    GlobalVariable.strBTableLine5 = "PCH5";
                    break;

                case 19:
                    GlobalVariable.strBDocType = "AP Debit Memo";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                    GlobalVariable.intBObjType = 19;
                    GlobalVariable.strBTableHeader = "ORPC";
                    GlobalVariable.strBTableLine1 = "RPC1";
                    GlobalVariable.strBTableLine3 = "RPC3";
                    GlobalVariable.strBTableLine5 = "RPC5";
                    break;

                case 20:
                    GlobalVariable.strBDocType = "Goods Receipt PO";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
                    GlobalVariable.intBObjType = 20;
                    GlobalVariable.strBTableHeader = "OPDN";
                    GlobalVariable.strBTableLine1 = "PDN1";
                    GlobalVariable.strBTableLine3 = "PDN3";
                    GlobalVariable.strBTableLine5 = "PDN5";
                    break;

                case 21:
                    GlobalVariable.strBDocType = "Goods Return";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
                    GlobalVariable.intBObjType = 21;
                    GlobalVariable.strBTableHeader = "ORPD";
                    GlobalVariable.strBTableLine1 = "RPD1";
                    GlobalVariable.strBTableLine3 = "RPD3";
                    GlobalVariable.strBTableLine5 = "RPD5";
                    break;

                case 22:
                    GlobalVariable.strBDocType = "Purchase Order";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                    GlobalVariable.intBObjType = 22;
                    GlobalVariable.strBTableHeader = "OPOR";
                    GlobalVariable.strBTableLine1 = "POR1";
                    GlobalVariable.strBTableLine3 = "POR3";
                    GlobalVariable.strBTableLine5 = "POR5";
                    break;

                case 23:
                    GlobalVariable.strBDocType = "Sales Quotations";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseQuotations;
                    GlobalVariable.strBTableHeader = "OQUT";
                    GlobalVariable.strBTableLine1 = "QUT1";
                    GlobalVariable.strBTableLine3 = "QUT3";
                    GlobalVariable.strBTableLine5 = "QUT5";
                    break;

                case 24:
                    GlobalVariable.strBDocType = "Incoming Payment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oIncomingPayments;
                    GlobalVariable.strBTableHeader = "ORCT";
                    GlobalVariable.strBTableLine1 = "RCT1";
                    GlobalVariable.strBTableLine2 = "RCT2";
                    GlobalVariable.strBTableLine3 = "RCT3";
                    break;

                case 30:
                    GlobalVariable.strBDocType = "Journal Entry";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oJournalEntries;
                    GlobalVariable.strBTableHeader = "OJDT";
                    GlobalVariable.strBTableLine1 = "JDT1";
                    break;

                case 46:
                    GlobalVariable.strBDocType = "Outgoing Payment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oVendorPayments;
                    GlobalVariable.strBTableHeader = "OVPM";
                    break;

                case 59:
                    GlobalVariable.strBDocType = "Goods Receipt";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
                    GlobalVariable.strBTableHeader = "OIGN";
                    GlobalVariable.strBTableLine1 = "IGN1";
                    break;

                case 60:
                    GlobalVariable.strBDocType = "Goods Issue";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;
                    GlobalVariable.strBTableHeader = "OIGE";
                    GlobalVariable.strBTableLine1 = "IGE1";
                    break;

                case 67:
                    GlobalVariable.strBDocType = "Stock Transfer";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                    GlobalVariable.strBTableHeader = "OWTR";
                    GlobalVariable.strBTableLine1 = "WTR1";
                    break;

                case 112:
                    GlobalVariable.strBDocType = "Draft";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oDrafts;
                    GlobalVariable.strBTableHeader = "ODRF";
                    GlobalVariable.strBTableLine1 = "DRF1";
                    break;

                case 204:
                    GlobalVariable.strBDocType = "AP DownPayment";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments;
                    GlobalVariable.strBTableHeader = "ODPO";
                    GlobalVariable.strBTableLine1 = "DPO1";
                    break;

                case 1250000001:
                    GlobalVariable.strBDocType = "Stock Transfer Request";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest;
                    GlobalVariable.strBTableHeader = "OWTQ";
                    GlobalVariable.strBTableLine1 = "WTQ1";
                    break;

                case    310:
                    GlobalVariable.strBDocType = "Fixed Asset Retirement";
                    GlobalVariable.intBObjType = 310;
                    //GlobalVariable.oBObjectType =
                    GlobalVariable.strBTableHeader = "ORTI";
                    GlobalVariable.strBTableLine1 = "RTI1";
                    break;

                case 1470000049:
                    GlobalVariable.strBDocType = "Fixed Asset Capitalization";
                    GlobalVariable.intBObjType = 1470000049;
                    //GlobalVariable.oBObjectType =
                    GlobalVariable.strBTableHeader = "OACQ";
                    GlobalVariable.strBTableLine1 = "ACQ1";
                    break;

                case 1470000113:
                    GlobalVariable.strBDocType = "Purchase Request";
                    GlobalVariable.oBObjectType = SAPbobsCOM.BoObjectTypes.oPurchaseRequest;
                    GlobalVariable.strBTableHeader = "OPRQ";
                    GlobalVariable.strBTableLine1 = "PRQ1";
                    break;

                default:
                    GlobalVariable.strBDocType = "";
                    GlobalVariable.oBObjectType = 0;
                    GlobalVariable.strBTableHeader = "";
                    break;

            }
        }
        public static void sendAlert(string strStatus, string strProcess, string strMsgTxt, SAPbobsCOM.BoObjectTypes ObjType, string strObjKey)
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Messages oMessages;

            string strSubject = "FT SAP B1 Services - " + strProcess;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT \"USER_CODE\", \"U_NAME\" FROM OUSR WHERE \"U_IntMsg\" = 'Y' ");

            if (oRecordset.RecordCount > 0)
            {
                oMessages = null;
                oMessages = (SAPbobsCOM.Messages)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                oMessages.Subject = strSubject;
                oMessages.MessageText = strMsgTxt;
                if (strStatus != "E")
                    oMessages.AddDataColumn("Document #", strObjKey, ObjType, strObjKey);
                oMessages.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;


                while (!(oRecordset.EoF))
                {
                    if (oRecordset.RecordCount > 1)
                        oMessages.Recipients.Add();

                    oMessages.Recipients.UserCode = oRecordset.Fields.Item("USER_CODE").Value.ToString();
                    oMessages.Recipients.NameTo = oRecordset.Fields.Item("U_NAME").Value.ToString();
                    oMessages.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    oMessages.Recipients.UserType = SAPbobsCOM.BoMsgRcpTypes.rt_InternalUser;

                    oRecordset.MoveNext();
                }

                if (oMessages.Add() != 0)
                {
                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                    SystemFunction.errorAppend(GlobalVariable.intErrNum.ToString() + " - " + GlobalVariable.strErrMsg);
                }
            }
        }

        public static string getCodebyId(string tableName, string strId, string ColumnName, string UDF)
        {
            string ObjectCode = string.Empty;

            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"SELECT {ColumnName} FROM {tableName} WHERE {UDF} = '{strId}'";
            oRecordset.DoQuery(query);

            ObjectCode = oRecordset.Fields.Item($"{ColumnName}").Value;

            return ObjectCode;
        }

        public static string getDocNum(int ObjType, string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT " + GlobalVariable.strTableHeader + ".\"DocNum\" FROM " + GlobalVariable.strTableHeader + " WHERE  " + GlobalVariable.strTableHeader + ".\"DocEntry\" = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }

        public static string getDocEntrybyRefNum(string strU_RefNum, string strTable)
        {
            string strDocEntry;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT DocEntry FROM " + strTable + " WHERE U_RefNum = '" + strU_RefNum + "' ");

            if (oRecordset.RecordCount > 0)
                strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();
            else
                strDocEntry = "0";

            return strDocEntry;
        }
        public static string getDocNumbyId(int ObjType, string strId)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT " + GlobalVariable.strTableHeader + ".\"DocNum\" FROM " + GlobalVariable.strTableHeader + " WHERE  " + GlobalVariable.strTableHeader + ".\"U_Id\" = '" + strId + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }

        public static string getU_RefNum(int ObjType, string strId)
        {
            string strU_RefNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT " + GlobalVariable.strTableHeader + ".\"U_RefNum\" FROM " + GlobalVariable.strTableHeader + " WHERE  " + GlobalVariable.strTableHeader + ".\"U_Id\" = '" + strId + "' ");

            if (oRecordset.RecordCount > 0)
                strU_RefNum = oRecordset.Fields.Item("U_RefNum").Value.ToString();
            else
                strU_RefNum = "0";

            return strU_RefNum;
        }
        public static string getJENum(string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT Number FROM OJDT WHERE TransID = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("Number").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }
        public static string getJVNum(string strDocEntry)
        {
            string strDocNum;

            SAPbobsCOM.Recordset oRecordset;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT BatchNum FROM OBTD WHERE BatchNum = '" + strDocEntry + "' ");

            if (oRecordset.RecordCount > 0)
                strDocNum = oRecordset.Fields.Item("Number").Value.ToString();
            else
                strDocNum = "0";

            return strDocNum;
        }
        public static void createResponse(string strFileName, string strStatus, string strATECSAPDoc, string strRemarks, string strDate, string strTime)
        {
            string strXMLPath, strALBSAPDoc;

            string[] strFValue;


            try
            {
                strFValue = strFileName.Split(Convert.ToChar("_"));
                strALBSAPDoc = strFValue[3];

                strXMLPath = GlobalVariable.strExpPath + @"\RES_" + strFileName;

                XmlTextWriter xWriter = new XmlTextWriter(strXMLPath, Encoding.UTF8);
                xWriter.Formatting = Formatting.Indented;

                xWriter.WriteStartElement("ResponseFile");

                xWriter.WriteStartElement("BaseDocNum");
                xWriter.WriteString(strALBSAPDoc);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Status");
                xWriter.WriteString(strStatus);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("TargetDocNum");
                xWriter.WriteString(strATECSAPDoc);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Remarks");
                xWriter.WriteString(strRemarks);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Date");
                xWriter.WriteString(strDate);
                xWriter.WriteEndElement();

                xWriter.WriteStartElement("Time");
                xWriter.WriteString(strTime);
                xWriter.WriteEndElement();

                xWriter.WriteEndElement();
                xWriter.Close();

            }
            catch (Exception ex)
            {

            }
        }
        public static string getSAPCode(string strCode, string strTable, string strFldCon1, string strFldVal1, string strAddCon)
        {
            string strRetCode, strQuery;
            SAPbobsCOM.Recordset oRecordset;

            if (!(string.IsNullOrEmpty(strAddCon)))
                strQuery = string.Format("SELECT \"{0}\" AS \"Code\" FROM {1} WHERE \"{2}\" = '{3}' AND {4} ", strCode, strTable, strFldCon1, strFldVal1, strAddCon);
            else
                strQuery = string.Format("SELECT \"{0}\" AS \"Code\"  FROM {1} WHERE \"{2}\" = '{3}' ", strCode, strTable, strFldCon1, strFldVal1);

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(strQuery);

            if (oRecordset.RecordCount > 0)
                strRetCode = oRecordset.Fields.Item("Code").Value.ToString();
            else
                strRetCode = "";

            SystemFunction.releaseObj(oRecordset);

            return strRetCode;
        }
        public static DateTime getDateTime(string strDateTime, string strOrigFormat, string strRetFormat)
        {
            DateTime dteRetDate = Convert.ToDateTime("01/01/9999");
            string strDateVal;

            if (!(string.IsNullOrEmpty(strDateTime)))
            {
                if (strRetFormat == "MM/DD/YYYY")
                {
                    if (strOrigFormat == "YYYYMMDD")
                    {

                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(4, 2)) > 12)
                                strDateVal = strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 2) + "/" + strDateTime.Substring(0, 4);
                            else
                                strDateVal = strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 1) + "/" + strDateTime.Substring(0, 4);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 1) + "/" + strDateTime.Substring(0, 4);
                        else
                            strDateVal = strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2) + "/" + strDateTime.Substring(0, 4);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                    else if (strOrigFormat == "MMDDYYYY")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(0, 2)) > 12)
                                strDateVal = strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 2) + "/" + strDateTime.Substring(3, 4);
                            else
                                strDateVal = strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 1) + "/" + strDateTime.Substring(3, 4);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 1) + "/" + strDateTime.Substring(2, 4);
                        else
                            strDateVal = strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 2) + "/" + strDateTime.Substring(4, 4);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                }
                else if (strRetFormat == "YYYY/MM/DD")
                {
                    if (strOrigFormat == "MMDDYYYY")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(0, 2)) > 12)
                                strDateVal = strDateTime.Substring(3, 4) + "/0" + strDateTime.Substring(0, 1) + "/" + strDateTime.Substring(1, 2);
                            else
                                strDateVal = strDateTime.Substring(3, 4) + "/" + strDateTime.Substring(0, 2) + "/" + strDateTime.Substring(2, 1);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(2, 4) + "/0" + strDateTime.Substring(0, 1) + "/0" + strDateTime.Substring(1, 1);
                        else
                            strDateVal = strDateTime.Substring(0, 4) + "/" + strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }
                    else if (strOrigFormat == "YYYYMMDD")
                    {
                        if (strDateTime.Length == 7)
                        {
                            if (Convert.ToInt32(strDateTime.Substring(4, 2)) > 12)
                                strDateVal = strDateTime.Substring(0, 4) + "/0" + strDateTime.Substring(4, 1) + "/" + strDateTime.Substring(5, 2);
                            else
                                strDateVal = strDateTime.Substring(0, 4) + strDateTime.Substring(4, 2) + "/0" + strDateTime.Substring(6, 1);
                        }
                        else if (strDateTime.Length == 6)
                            strDateVal = strDateTime.Substring(0, 4) + "/0" + strDateTime.Substring(4, 1) + "/0" + strDateTime.Substring(5, 1);
                        else
                            strDateVal = strDateTime.Substring(0, 4) + "/" + strDateTime.Substring(4, 2) + "/" + strDateTime.Substring(6, 2);

                        dteRetDate = Convert.ToDateTime(strDateVal);
                    }

                }

            }
            else
                dteRetDate = Convert.ToDateTime("01/01/1900");

            return dteRetDate;


        }
        public static bool importXLSX(string strXLSPath, string strHeader, string strSheet)
        {
            try
            {
                string strConnString;

                GlobalVariable.oDTImpData = new DataTable("ImportData");
                GlobalVariable.oDTImpData.Reset();

                if (Path.GetExtension(strXLSPath) == ".xlsx")
                    strConnString = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0; HDR={1};' ", strXLSPath, strHeader);
                else
                    strConnString = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0; HDR={1};' ", strXLSPath, strHeader);

                OleDbConnection oledbConn = new OleDbConnection(strConnString);

                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}$]", strSheet), oledbConn);

                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                oleda.Fill(GlobalVariable.oDTImpData);

                oledbConn.Close();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.errorAppend(string.Format("Error retrieving data from Excel ({0} - {1}). Description : {2} ", strXLSPath, strSheet, ex.Message.ToString()));

                return false;
            }

        }
        public static bool importCSV(string strFilePath, string strCSVPath, string strHeader, string strDlmtd)
        {
            try
            {

                GlobalVariable.oDTImpData = new DataTable("ImportData");
                GlobalVariable.oDTImpData.Reset();

                string connString = string.Format("Provider = Microsoft.Jet.OleDb.4.0; Data Source = {0}; Extended Properties = 'text; HDR = {1}; FMT = Delimited ({2})' ", strFilePath, strHeader, strDlmtd);

                OleDbConnection oledbConn = new OleDbConnection(connString);

                oledbConn.Open();

                OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", strCSVPath), oledbConn);

                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                oleda.Fill(GlobalVariable.oDTImpData);

                oledbConn.Close();

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.errorAppend(string.Format("Error retrieving data from CSV File ({0}). Description : {1} ", strCSVPath, ex.Message.ToString()));

                return false;
            }

        }
        public static bool decryptPGP(string strEncrytdFilePath, string strPrivateKeyPath, string strPassKey, string strDcryptdFilePath)
        {


            try
            {
                // Load keys
                FileInfo privateKey = new FileInfo(strPrivateKeyPath);
                EncryptionKeys encryptionKeys = new EncryptionKeys(privateKey, strPassKey);

                // Reference input/output files
                FileInfo inputFile = new FileInfo(strEncrytdFilePath);
                FileInfo decryptedFile = new FileInfo(strDcryptdFilePath);

                // Decrypt
                PGP pgp = new PGP(encryptionKeys);
                pgp.DecryptFileAsync(inputFile, decryptedFile);

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                return false;
            }
        }
        public static void cleanTempFiles()
        {
            DateTime dteStart = DateTime.Now;

            try
            {
                foreach (var strFile in Directory.GetFiles(GlobalVariable.strTempPath, "*.*"))
                {
                    if (File.Exists(strFile))
                        File.Delete(strFile);
                }
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "Clean Temporary Files", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }

        public static bool checkRefNum(string strU_RefNum, string strTableHeader)
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Validation: Check if U_RefNum exists in Table Header
            string validationQuery = $"SELECT COUNT(*) AS RecordCount FROM {strTableHeader} WHERE U_RefNum = '{strU_RefNum}'";
            oRecordset.DoQuery(validationQuery);

            int recordCount = (int)oRecordset.Fields.Item("RecordCount").Value;

            if (recordCount > 0)
                return false;
            else
                return true;
        }

        public static DateTime stringToDateFormat(string strDate)
        {
            string format = "yyyyMMdd";
            DateTime dteDate;
            DateTime.TryParseExact(strDate, format, null, System.Globalization.DateTimeStyles.None, out dteDate);

            return dteDate;
        }

        public static string getAccountCode(string strId)
        {
            string ObjectCode = string.Empty;
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = $"SELECT AcctCode FROM OACT WHERE FormatCode = '{strId}'";
            oRecordset.DoQuery(query);
            ObjectCode = oRecordset.Fields.Item($"AcctCode").Value;
            return ObjectCode;
        }
    }
}
