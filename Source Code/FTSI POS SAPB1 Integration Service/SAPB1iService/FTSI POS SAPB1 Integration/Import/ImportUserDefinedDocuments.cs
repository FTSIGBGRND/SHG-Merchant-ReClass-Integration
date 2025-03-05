using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using ClosedXML.Excel;

namespace SAPB1iService
{
    class ImportUserDefinedDocuments
    {
        private static DateTime dteStart;
        private static string strTransType;

        private static string strMsgBod, strStatus, strPostDocNum, strPostDocEnt;

        private static DataTable oDTPayment;

        private static int intSeries;
        public static void _ImportUserDefinedDocuments()
        {

            SAPbobsCOM.Recordset oRecordset;

            strTransType = "Documents - Credit Card Auto AR ReClass";

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery("SELECT NNM1.\"Series\" " +
                               "FROM \"NNM1\" INNER JOIN \"OADM\" ON NNM1.\"SeriesName\" = OADM.\"U_InvSeries\" " +
                               "WHERE NNM1.\"ObjectCode\" = '13' ");

            if (oRecordset.RecordCount > 0)
                intSeries = Convert.ToInt32(oRecordset.Fields.Item("Series").Value.ToString());
            else
            {
                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", "-999", "Please setup AR Invoice Numbering Series for AR Merchant ReClass in Administration > Company Details.");
                return;
            }

            importCCAutoARReClass();
        }
        private static void importCCAutoARReClass()
        {
            string strQuery;

            int intRow;

            SAPbobsCOM.Recordset oRecordset;

            try
            {
                initDTPayment();

                strTransType = "Documents - Credit Card Auto AR ReClass";

                if (GlobalVariable.strDBType != "HANA DB")
                    strQuery = string.Format("EXEC FTSI_B1IS_IMPORT_POSINTEGRATION_CCAUTOARRECLASS");
                else
                    strQuery = string.Format("CALL \"FTSI_B1IS_IMPORT_POSINTEGRATION_CCAUTOARRECLASS\" ");

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    intRow = 0;

                    oRecordset.MoveFirst();

                    while (!oRecordset.EoF)
                    {
                        oDTPayment.Rows.Add(intRow,
                                            oRecordset.Fields.Item("DocEntry").Value.ToString(),
                                            oRecordset.Fields.Item("DocNum").Value.ToString(),
                                            oRecordset.Fields.Item("CardCode").Value.ToString(),
                                            Convert.ToDateTime(oRecordset.Fields.Item("DocDate").Value.ToString()),
                                            oRecordset.Fields.Item("CreditCard").Value.ToString(),
                                            oRecordset.Fields.Item("CrCardNum").Value.ToString(),
                                            oRecordset.Fields.Item("CardName").Value.ToString(),
                                            oRecordset.Fields.Item("AcctCode").Value.ToString(),
                                            oRecordset.Fields.Item("U_RCVatGroup").Value.ToString(),
                                            Convert.ToDouble(oRecordset.Fields.Item("CreditSum").Value.ToString()),
                                            oRecordset.Fields.Item("U_CardCode").Value.ToString(),
                                            oRecordset.Fields.Item("U_BCGLAcct").Value.ToString(),
                                            oRecordset.Fields.Item("U_BCVatGroup").Value.ToString(),
                                            Convert.ToDouble(oRecordset.Fields.Item("U_BCRate").Value.ToString()),
                                            oRecordset.Fields.Item("U_WTCode").Value.ToString());

                        oRecordset.MoveNext();

                        intRow++;

                    }

                    postCCAutoARReClass();
                }

            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();
            }
        }
        private static void postCCAutoARReClass()
        {
            DataTable oDTIncoming, oDTHeader, oDTDetails;

            string strDocEntry = "", strDocNum = "", strBCardCode, strTCardCode, strWTCode, 
                   strBCGLAcct, strBCCGLAcct, strCreditCard, strCrCardName, strCompany;

            double dblRate, dblBnkChrg, dblCrediSum;

            DateTime dteDoc;

            SAPbobsCOM.Documents oDocuments;
            SAPbobsCOM.Recordset oRecordset;

            bool blWithErr = false;

            try
            {
                GlobalFunction.getObjType(13);

                if (oDTPayment.Rows.Count > 0)
                {

                    oDTIncoming = oDTPayment.DefaultView.ToTable(true, "DocEntry", "DocNum", "BCardCode", "DocDate");

                    if (oDTIncoming.Rows.Count > 0)
                    {
                        for (int intRowE = 0; intRowE <= oDTIncoming.Rows.Count - 1; intRowE++)
                        {
                            blWithErr = false;

                            if (!(GlobalVariable.oCompany.InTransaction))
                                GlobalVariable.oCompany.StartTransaction();

                            strDocEntry = oDTIncoming.Rows[intRowE]["DocEntry"].ToString();
                            strDocNum = oDTIncoming.Rows[intRowE]["DocNum"].ToString();
                            dteDoc = Convert.ToDateTime(oDTIncoming.Rows[intRowE]["DocDate"].ToString());
                            strBCardCode = oDTIncoming.Rows[intRowE]["BCardCode"].ToString();

                            oDTHeader = oDTPayment.Select("DocEntry = '" + strDocEntry + "' ").CopyToDataTable().DefaultView.ToTable(true, "TCardCode", "WTCode", "CreditCard", "CrCardNam");

                            if (oDTHeader.Rows.Count > 0)
                            {
                                for (int intRowH = 0; intRowH <= oDTHeader.Rows.Count - 1; intRowH++)
                                {
                                    strTCardCode = oDTHeader.Rows[intRowH]["TCardCode"].ToString();
                                    strWTCode = oDTHeader.Rows[intRowH]["WTCode"].ToString();
                                    strCreditCard = oDTHeader.Rows[intRowH]["CreditCard"].ToString();
                                    strCrCardName = oDTHeader.Rows[intRowH]["CrCardNam"].ToString();

                                    oRecordset = null;
                                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    oRecordset.DoQuery(string.Format("SELECT \"CardCode\" FROM OCRD WHERE \"CardCode\" = '{0}'", strTCardCode));

                                    if (!(oRecordset.RecordCount > 0))
                                    {
                                        blWithErr = true;

                                        strStatus = "E";
                                        GlobalVariable.intErrNum = 899;
                                        GlobalVariable.strErrMsg = string.Format("Invalid Business Partner {0}. Please Check Credit Card Setup.", strTCardCode);

                                        strMsgBod = string.Format("Error Posting {0} from Incoming Payment {1} - {2}.\r" +
                                                                  "Error Code: {3}\rDescription: {4} ", strTransType, strDocNum, strBCardCode,
                                                                                                        GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strDocNum, "", "", dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        updateBaseDoc(strStatus, strDocEntry, strDocNum);
                                        
                                        GlobalFunction.sendAlert(strStatus, strTransType, strMsgBod, GlobalVariable.oObjectType, "0");

                                        break;
                                    }

                                    oDocuments = null;
                                    oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                    oDocuments.Series = intSeries;

                                    oDocuments.CardCode = strTCardCode;
                                    oDocuments.DocDate = dteDoc;
                                    oDocuments.TaxDate = dteDoc;
                                    oDocuments.DocDueDate = dteDoc;

                                    oDocuments.Comments = string.Format("Incoming Payment {0}_{1}_{2}_{3}.", strDocNum, strBCardCode, strCrCardName, dteDoc.ToString("MMddyyyy"));

                                    oDocuments.DocType = BoDocumentTypes.dDocument_Service;

                                    oDocuments.UserFields.Fields.Item("U_ORCTDocEnt").Value = strDocEntry;
                                    oDocuments.UserFields.Fields.Item("U_ORCTDocNum").Value = strDocNum;

                                    oDTDetails = oDTPayment.Select("DocEntry = '" + strDocEntry + "' AND CreditCard = '" + strCreditCard + "' AND TCardCode = '" + strTCardCode + "' AND WTCode = '" + strWTCode + "' ").CopyToDataTable().DefaultView.ToTable();

                                    if (oDTDetails.Rows.Count > 0)
                                    {
                                        for (int intRowD = 0; intRowD <= oDTDetails.Rows.Count - 1; intRowD++)
                                        {

                                            strBCCGLAcct = oDTDetails.Rows[intRowD]["BCGLAcct"].ToString();
                                            dblCrediSum = Convert.ToDouble(oDTDetails.Rows[intRowD]["CreditSum"].ToString());
                                            dblRate = Convert.ToDouble(oDTDetails.Rows[intRowD]["BCRate"].ToString());

                                            dblBnkChrg = Math.Round((dblCrediSum * dblRate / 100), 2) * -1;

                                            strBCGLAcct = GlobalFunction.getSAPCode("AcctCode", "OACT", "FormatCode", strBCCGLAcct.Replace("-", ""), "");
                                            if (string.IsNullOrEmpty(strBCGLAcct))
                                            {
                                                blWithErr = true;

                                                strStatus = "E";
                                                GlobalVariable.intErrNum = 898;
                                                GlobalVariable.strErrMsg = string.Format("Invalid Bank Charge GL Account {0} for Credit Card {1} - {2}. Please check credit card setup.", strBCCGLAcct, strCreditCard, strCrCardName);

                                                strMsgBod = string.Format("Error Posting {0} from Incoming Payment {1} - {2}.\r" +
                                                                          "Error Code: {3}\rDescription: {4} ", strTransType, strDocNum, strBCardCode, 
                                                                                                                GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strDocNum, "", "", dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                                break;
                                            }

                                            if (intRowD > 0)
                                                oDocuments.Lines.Add();

                                            oDocuments.Lines.AccountCode = oDTDetails.Rows[intRowD]["AcctCode"].ToString();
                                            oDocuments.Lines.ItemDescription = "Payment - " + strCrCardName;
                                            oDocuments.Lines.UnitPrice = dblCrediSum;
                                            oDocuments.Lines.DiscountPercent = 0;
                                            oDocuments.Lines.VatGroup = oDTDetails.Rows[intRowD]["RCVatGroup"].ToString();
                                            oDocuments.Lines.WTLiable = (string.IsNullOrEmpty(strWTCode)) ? BoYesNoEnum.tNO : BoYesNoEnum.tYES;

                                            //additional request and customized cost center per database 12/13/2022

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                            oRecordset.DoQuery(string.Format("SELECT \"AliasName\" FROM OADM "));

                                            strCompany = oRecordset.Fields.Item("AliasName").Value.ToString();

                                            if (strCompany == "NKWM")
                                            {
                                                oDocuments.Lines.CostingCode2 = strBCardCode;
                                                oDocuments.Lines.CostingCode5 = "NFOP01";
                                            }
                                            else
                                            {
                                                oDocuments.Lines.CostingCode = strBCardCode;
                                                oDocuments.Lines.CostingCode2 = "FOP01";
                                            }

                                            oDocuments.Lines.Add();
                                            oDocuments.Lines.AccountCode = strBCGLAcct;
                                            oDocuments.Lines.ItemDescription = "Bank Charge - " + strCrCardName;
                                            oDocuments.Lines.UnitPrice = dblBnkChrg;
                                            oDocuments.Lines.DiscountPercent = 0;
                                            oDocuments.Lines.VatGroup = oDTDetails.Rows[intRowD]["BCVatGroup"].ToString();
                                            oDocuments.Lines.WTLiable = BoYesNoEnum.tNO;

                                            if (strCompany == "NKWM")
                                            {
                                                oDocuments.Lines.CostingCode2 = strBCardCode;
                                                oDocuments.Lines.CostingCode5 = "NFOP01";
                                            }
                                            else
                                            {
                                                oDocuments.Lines.CostingCode = strBCardCode;
                                                oDocuments.Lines.CostingCode2 = "FOP01";
                                            }
                                        }
                                    }

                                    if (blWithErr == false)
                                    {
                                        if (!(string.IsNullOrEmpty(strWTCode)))
                                            oDocuments.WithholdingTaxData.WTCode = strWTCode;
                                        
                                        if (oDocuments.Add() != 0)
                                        {
                                            blWithErr = true;

                                            strStatus = "E";
                                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                            strMsgBod = string.Format("Error Posting {0} from Incoming Payment {1} - {2}.\r" +
                                                                      "Error Code: {3}\r" +
                                                                      "Description: {4} ", strTransType, strDocNum, strBCardCode,
                                                                                           GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strDocNum, "", "", dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                            updateBaseDoc(strStatus, strDocEntry, strDocNum);

                                            break;
                                        }
                                        else
                                        {

                                            strStatus = "S";

                                            strPostDocEnt = GlobalVariable.oCompany.GetNewObjectKey().ToString();
                                            strPostDocNum = GlobalFunction.getDocNum(GlobalVariable.intObjType, strPostDocEnt);

                                            GlobalVariable.intErrNum = 0;
                                            GlobalVariable.strErrMsg = string.Format("Successfully Posted {0} with Document No {1} from Incoming Payment {2} - {3}. ", strTransType, strPostDocNum, strDocNum, strBCardCode);

                                            strMsgBod = string.Format("Successfully Posted {0} from Incoming Payment {1} - {2}. ", strTransType, strDocNum, strBCardCode);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strDocNum, strPostDocEnt, strPostDocNum, dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                        }
                                    }
                                    else
                                        break;

                                }
                            }

                            updateBaseDoc(strStatus, strDocEntry, strDocNum);

                            GlobalFunction.sendAlert(strStatus, strTransType, strMsgBod, GlobalVariable.oObjectType, "0");

                            if (GlobalVariable.oCompany.InTransaction)
                                GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                strStatus = "E";

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strDocNum, "", "", dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                if (!(string.IsNullOrEmpty(strDocEntry)))
                    updateBaseDoc(strStatus, strDocEntry, strDocNum);
            }
        }

        private static bool updateBaseDoc(string strStatus, string strBDocEntry, string strBDocNum)
        {
            string strQuery;

            try
            {

                strQuery = string.Format("UPDATE ORCT SET \"U_isExtract\" = '{0}' " +
                                         "WHERE \"DocEntry\" = '{1}' AND \"DocNum\" = '{2}' ", strStatus, strBDocEntry, strBDocNum);
                
                if (!(SystemFunction.executeQuery(strQuery)))
                {
                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strBDocNum, "", "", dteStart, "E", "-999", "Error Updating Incoming Payment Base Document!");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), strBDocNum, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;

            }
        }
        private static void initDTPayment()
        {
            oDTPayment = new DataTable("CC Payment");
            oDTPayment.Columns.Add("Row", typeof(System.Int32));
            oDTPayment.Columns.Add("DocEntry", typeof(System.String));
            oDTPayment.Columns.Add("DocNum", typeof(System.String));
            oDTPayment.Columns.Add("BCardCode", typeof(System.String));
            oDTPayment.Columns.Add("DocDate", typeof(System.DateTime));
            oDTPayment.Columns.Add("CreditCard", typeof(System.String));
            oDTPayment.Columns.Add("CrCardNum", typeof(System.String));
            oDTPayment.Columns.Add("CrCardNam", typeof(System.String));
            oDTPayment.Columns.Add("AcctCode", typeof(System.String));
            oDTPayment.Columns.Add("RCVatGroup", typeof(System.String));
            oDTPayment.Columns.Add("CreditSum", typeof(System.Double));
            oDTPayment.Columns.Add("TCardCode", typeof(System.String));
            oDTPayment.Columns.Add("BCGLAcct", typeof(System.String));
            oDTPayment.Columns.Add("BCVatGroup", typeof(System.String));
            oDTPayment.Columns.Add("BCRate", typeof(System.String));
            oDTPayment.Columns.Add("WTCode", typeof(System.String));
            oDTPayment.Clear();

        }
    }

}
