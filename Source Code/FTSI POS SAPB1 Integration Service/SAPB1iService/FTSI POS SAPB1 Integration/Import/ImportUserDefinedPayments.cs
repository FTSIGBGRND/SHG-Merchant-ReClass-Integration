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

namespace SAPB1iService
{
    class ImportUserDefinedPayments
    {
        private static DateTime dteStart;
        private static string strTransType = "Payment - IEMOP Incoming";
        private static string strMsgBod;
        public static void _ImportUserDefinedPayments()
        {
            importFromFile();
        }
        public static bool importFromObject(int intObjType, string strAPDocEntry, DateTime dtePayDate)
        {
            string strQuery, strCardCode, strStatus, strMsgBod;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Payments oPayments;

            try
            {
                strQuery = string.Format("SELECT OPCH.\"DocEntry\", OPCH.\"DocNum\", OPCH.\"CardCode\", OPCH.\"Comments\", OPCH.\"DocTotal\" " +
                                         "FROM OPCH " +
                                         "WHERE OPCH.\"DocEntry\" = '{0}' ", strAPDocEntry);

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {

                    strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();

                    GlobalFunction.getObjType(intObjType);

                    oPayments = null;
                    oPayments = (SAPbobsCOM.Payments)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                    oPayments.CardCode = strCardCode;
                    oPayments.DocDate = dtePayDate;
                    oPayments.DueDate = dtePayDate;
                    oPayments.TaxDate = dtePayDate;
                    oPayments.Remarks = oRecordset.Fields.Item("Comments").Value.ToString();

                    oPayments.TransferDate = dtePayDate;
                    oPayments.TransferSum = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                    oPayments.Invoices.DocEntry = Convert.ToInt32(strAPDocEntry);
                    oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_PurchaseInvoice;
                    oPayments.Invoices.SumApplied = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                    if (oPayments.Add() != 0)
                    {
                        GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();
                        
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = string.Format("Error Processing Outgoing Payment. {0}.", ex.Message.ToString());

                return false;
            }
        }
        private static void importFromFile()
        {

            string strStatus = "";

            try
            {
                string[] strFileImport = GlobalVariable.strImpExt.Split(Convert.ToChar("|"));

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, string.Format("*{0}", fileimport)))
                    {
                        GlobalVariable.strFileName = Path.GetFileName(strFile);

                        if (strFile.Contains("Incoming"))
                        {
                            dteStart = DateTime.Now;

                            if (importDIAPIPostPaymentFExcel(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";

                            TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));

                            GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                            //EmailSender._EmailSender("Import", strStatus, GlobalVariable.strFileName, strPostDocNum, string.Format("Error Code : {0} Description : {1} ", GlobalVariable.intErrNum, GlobalVariable.strErrMsg));
                        }
                    }
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", strTransType, "28", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static bool importDIAPIPostPaymentFExcel(string strFile)
        {
            string strSTLID = "", strNumAtCard = "", strCardCode, strTrnsfrGL, strQuery;

            bool blBPExist, blTempErr;

            double dblTotPymnt, dblDocTotal, dblPayAmnt = 0, dblRunBal;

            int intRowInv;

            DataTable oDTSTLID, oDTPayment;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Payments oPayments;

            DateTime dteDoc = Convert.ToDateTime("01/01/1900");
            DateTime dteDue = Convert.ToDateTime("01/01/1900");
            DateTime dteTax = Convert.ToDateTime("01/01/1900");
            DateTime dteTrnsfr = Convert.ToDateTime("01/01/1900");

            try
            {

                if (GlobalFunction.importXLSX(Path.GetFullPath(strFile), "NO", "Sheet1"))
                {
                    if (GlobalVariable.oDTImpData.Rows.Count > 0)
                    {
                        blBPExist = true;
                        blTempErr = false;

                        strQuery = string.Format("SELECT IEMOPGL.\"U_PymntGL\" " +
                                                 "FROM \"@FTIEMOPGL\" IEMOPGL " +
                                                 "WHERE IEMOPGL.\"Code\" = 'AR' ");
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (!(oRecordset.RecordCount > 0))
                        {
                            GlobalVariable.intErrNum = -998;
                            GlobalVariable.strErrMsg = "Please setup Transfer GL Account for Incoming Payment.";

                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }
                        else
                            strTrnsfrGL = oRecordset.Fields.Item("U_PymntGL").Value.ToString();

                        for (int intRowH = 0; intRowH <= 6; intRowH++)
                        {
                            strSTLID = GlobalVariable.oDTImpData.Rows[intRowH][0].ToString();

                            if (strSTLID == "Posting Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteDoc = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Due Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteDue = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Document Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteTax = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }

                            if (strSTLID == "Transfer Date")
                            {
                                if (validateDates(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString()))
                                    dteTrnsfr = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRowH][1].ToString());
                                else
                                    blTempErr = true;
                            }
                        }

                        if (blTempErr == true)
                        {
                            GlobalVariable.intErrNum = -998;
                            GlobalVariable.strErrMsg = "Template Error. Please check dates.";

                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }
                        
                        oDTSTLID = GlobalVariable.oDTImpData.Select(string.Format("{0} IS NOT NULL AND {0} <> 'Received From (Buyer STL ID)' ", GlobalVariable.oDTImpData.Columns[2].ColumnName)).CopyToDataTable().DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[2].ColumnName, GlobalVariable.oDTImpData.Columns[4].ColumnName);

                        if (!(GlobalVariable.oCompany.InTransaction))
                            GlobalVariable.oCompany.StartTransaction();

                        for (int intRowID = 0; intRowID <= oDTSTLID.Rows.Count - 1; intRowID++)
                        {
                            strSTLID = oDTSTLID.Rows[intRowID][0].ToString();
                            strNumAtCard = oDTSTLID.Rows[intRowID][1].ToString();

                            strQuery = string.Format("SELECT OCRD.\"CardCode\" " +
                                                        "FROM \"@FTIEMOP1\"  IEMOP1 INNER JOIN \"@FTOIEMOP\" OIEMOP ON IEMOP1.\"Code\" = OIEMOP.\"Code\"  " +
                                                        "                           INNER JOIN OCRD ON OIEMOP.\"Code\" = OCRD.\"CardCode\" " +
                                                        "WHERE IEMOP1.\"U_STLID\" = '{0}' AND OCRD.\"CardType\" = 'C' ", strSTLID);
                            oRecordset = null;
                            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            oRecordset.DoQuery(strQuery);

                            if (!(oRecordset.RecordCount > 0))
                            {
                                blBPExist = false;

                                //to do create excel template for business partner

                                continue;
                            }
                            else
                                strCardCode = oRecordset.Fields.Item("CardCode").Value.ToString();

                            if (blBPExist == true)
                            {

                                GlobalFunction.getObjType(24);

                                oPayments = null;
                                oPayments = (SAPbobsCOM.Payments)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                oPayments.CardCode = strCardCode;
                                oPayments.DocDate = dteDoc;
                                oPayments.DueDate = dteDue;
                                oPayments.TaxDate = dteTax;
                                
                                strQuery = string.Format("{0} = '{1}' AND {2} = '{3}' ", GlobalVariable.oDTImpData.Columns[2].ColumnName, strSTLID, GlobalVariable.oDTImpData.Columns[4].ColumnName, strNumAtCard);
                                oDTPayment = GlobalVariable.oDTImpData.Select(strQuery).CopyToDataTable().DefaultView.ToTable();
                                
                                dblTotPymnt = 0;

                                for (int intRowDP = 0; intRowDP <= oDTPayment.Rows.Count - 1; intRowDP++)
                                    dblTotPymnt = dblTotPymnt + Math.Abs(Convert.ToDouble(oDTPayment.Rows[intRowDP][10].ToString()));

                                oPayments.TransferAccount = strTrnsfrGL;
                                oPayments.TransferSum = Math.Abs(dblTotPymnt);
                                oPayments.TransferDate = dteTrnsfr;

                                strQuery = string.Format("SELECT OINV.\"DocEntry\", OINV.\"DocNum\", (OINV.\"DocTotal\" - OINV.\"PaidToDate\") AS \"DocTotal\" " +
                                                         "FROM OINV " +
                                                         "WHERE OINV.\"DocStatus\" = 'O' AND OINV.\"NumAtCard\" = '{0}' AND CardCode = '{1}' ", strNumAtCard, strCardCode);

                                oRecordset = null; 
                                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                oRecordset.DoQuery(strQuery);

                                if (oRecordset.RecordCount > 0)
                                {
                                    intRowInv = 0;
                                    dblRunBal = dblTotPymnt;

                                    while (!(oRecordset.EoF))
                                    {
                                        dblDocTotal = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString());

                                        if (dblRunBal > 0)
                                        {
                                            if (dblRunBal >= dblDocTotal)
                                                dblPayAmnt = dblDocTotal;
                                            else
                                                dblPayAmnt = dblRunBal;

                                            if (intRowInv > 0)
                                                oPayments.Invoices.Add();

                                            oPayments.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                            oPayments.Invoices.DocEntry = Convert.ToInt32(oRecordset.Fields.Item("DocEntry").Value.ToString());
                                            oPayments.Invoices.SumApplied = dblPayAmnt;

                                            dblRunBal = dblRunBal - dblPayAmnt;
                                        }
                                        else
                                            break;

                                        intRowInv++;

                                        oRecordset.MoveNext();
                                    }
                                }
                                else
                                {
                                    GlobalVariable.intErrNum = -997;
                                    GlobalVariable.strErrMsg = string.Format("AR Invoice Document No Found for {0} - {1}. ", strSTLID, strNumAtCard);

                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }

                                if (oPayments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", strTransType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }
                            }
                        }
                    }

                    strMsgBod = string.Format("Successfully Posted {0} - {1}.", strTransType, GlobalVariable.strFileName);

                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "S", "0", strMsgBod);

                    if (GlobalVariable.oCompany.InTransaction)
                        GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                    GC.Collect();

                    return true;
                }
                else
                    return false;

                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
        private static bool validateDates(string strDate)
        {
            DateTime dteReturn;

            if (!DateTime.TryParse(strDate, out dteReturn))
                return false;
            else
            {
                if (dteReturn == Convert.ToDateTime("01/01/1900"))
                    return false;
                else
                    return true;
            }

        }
    }
}
