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
    class ImportUserDefinedJournalVoucher
    {
        private static DateTime dteStart;
        private static string strTransType = "Journal Voucher - Payroll Integration";
        private static string strPostDocNum;
        private static string strMsgBod;
        private static string strCardCode;

        public static void _ImportUserDefinedJournalVoucher()
        {
            importFromFile();
        }
        private static void importFromFile()
        {
            string strStatus;

            try
            {                
                string[] strFileImport = new string[] { string.Format("*.pgp") };

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, fileimport))
                    {

                        dteStart = DateTime.Now;

                        string strPrivateKeyPath = GlobalVariable.strFilePath + @"\PGP Files\SAPSFCPIPGPKEY2_SECRET.asc";
                        string strPassKey = "Tdcxkey@2";
                        string strDcryptdFilePath = GlobalVariable.strTempPath +  Path.GetFileName(strFile).Replace(".csv.pgp", ".csv");

                        if (GlobalFunction.decryptPGP(strFile, strPrivateKeyPath, strPassKey, strDcryptdFilePath))
                        {
                            GlobalFunction.getObjType(28);
                            GlobalVariable.strFileName = Path.GetFileName(strDcryptdFilePath);

                            if (importDIAPIPostJV(GlobalVariable.strTempPath, strDcryptdFilePath))
                                strStatus = "S";
                            else
                                strStatus = "E";

                            TransferFile.transferArcFiles("Import", strDcryptdFilePath);

                            if (File.Exists(strDcryptdFilePath))
                                File.Delete(strDcryptdFilePath);
                        }
                        else
                        {
                            strStatus = "E";
                            strMsgBod = string.Format("Error Decrpting File. Error Code : {0} Description : {1} ", GlobalVariable.intErrNum, GlobalVariable.strErrMsg);                          
                        }

                        TransferFile.exportSFTPReturnFiles(strFile, strStatus);

                        TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));

                        GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                        EmailSender._EmailSender("Import", strStatus, GlobalVariable.strFileName, strPostDocNum, string.Format("Error Code : {0} Description : {1} ", GlobalVariable.intErrNum, GlobalVariable.strErrMsg));

                    }
                }

                //importDIAPIPostJV(GlobalVariable.strTempPath, GlobalVariable.strTempPath + "PHL_190722_121632.csv");

                GC.Collect();
            }
            catch(Exception ex)
            {
                SystemFunction.transHandler("Import", strTransType, "28", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static bool importDIAPIPostJV(string strFilePath, string strFile)
        {
            string strAcctCode, strEmpID, strStatus, strRefDate = "01/01/1900";
            
            string[] strRemarks = new string[] {} ;
            
            double dblDebit, dblCredit;

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.JournalVouchers oJournalVouchers;

            try
            {
                if (GlobalFunction.importCSV(strFilePath, Path.GetFileName(strFile), "YES", GlobalVariable.chrDlmtr.ToString()))
                {

                    if (!(GlobalVariable.oCompany.InTransaction))
                        GlobalVariable.oCompany.StartTransaction();

                    oJournalVouchers = null;
                    oJournalVouchers = (SAPbobsCOM.JournalVouchers)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oJournalVouchers);

                    for (int intRow = 0; intRow <= GlobalVariable.oDTImpData.Rows.Count - 1; intRow++)
                    {

                        if (intRow > 0)
                            oJournalVouchers.JournalEntries.Lines.Add();

                        strAcctCode = GlobalVariable.oDTImpData.Rows[intRow][2].ToString();
                        
                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(string.Format("SELECT \"LocManTran\" FROM OACT WHERE \"AcctCode\" = '{0}' ", strAcctCode));

                        if (oRecordset.RecordCount > 0)
                        {
                            if (oRecordset.Fields.Item("LocManTran").Value.ToString() == "Y")
                            {

                                if (importDIAPIPostBP(GlobalVariable.oDTImpData.Rows[intRow][0].ToString(), GlobalVariable.oDTImpData.Rows[intRow][1].ToString()))
                                    oJournalVouchers.JournalEntries.Lines.ShortName = strCardCode;
                                else
                                {
                                    strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }
                            }
                            else
                                oJournalVouchers.JournalEntries.Lines.AccountCode = GlobalVariable.oDTImpData.Rows[intRow][2].ToString();
                        }
                        else
                        {
                            GlobalVariable.intErrNum =-999;
                            GlobalVariable.strErrMsg = string.Format("GL Account '{0}' not exist in SAP Business One at line {1}. ", strAcctCode, intRow + 2);

                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            return false;
                        }

                        if (!(string.IsNullOrEmpty(GlobalVariable.oDTImpData.Rows[intRow][10].ToString())))
                            dblDebit = Convert.ToDouble(GlobalVariable.oDTImpData.Rows[intRow][10].ToString());
                        else
                            dblDebit = 0;

                        if (!(string.IsNullOrEmpty(GlobalVariable.oDTImpData.Rows[intRow][11].ToString())))
                            dblCredit = Convert.ToDouble(GlobalVariable.oDTImpData.Rows[intRow][11].ToString());
                        else
                            dblCredit = 0;

                        oJournalVouchers.JournalEntries.Lines.Debit = dblDebit;
                        oJournalVouchers.JournalEntries.Lines.Credit = dblCredit;

                        oJournalVouchers.JournalEntries.Lines.ProjectCode = GlobalVariable.oDTImpData.Rows[intRow][5].ToString();
                        oJournalVouchers.JournalEntries.Lines.CostingCode = GlobalVariable.oDTImpData.Rows[intRow][6].ToString();
                        oJournalVouchers.JournalEntries.Lines.CostingCode4 = GlobalVariable.oDTImpData.Rows[intRow][7].ToString();
                        oJournalVouchers.JournalEntries.Lines.CostingCode2 = GlobalVariable.oDTImpData.Rows[intRow][8].ToString();

                        oJournalVouchers.JournalEntries.Lines.LineMemo = GlobalVariable.oDTImpData.Rows[intRow][12].ToString();

                        if (intRow == 0)
                        {
                            strRefDate = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[intRow][14].ToString()).ToString("MM/dd/yyyy");
                            strRemarks = GlobalVariable.oDTImpData.Rows[intRow][12].ToString().Split(Convert.ToChar("_"));
                        }
                            
                            
                    }

                    oJournalVouchers.JournalEntries.ReferenceDate = Convert.ToDateTime(strRefDate);
                    oJournalVouchers.JournalEntries.DueDate = Convert.ToDateTime(strRefDate); 
                    oJournalVouchers.JournalEntries.TaxDate = Convert.ToDateTime(strRefDate);
                    oJournalVouchers.JournalEntries.Memo = strRemarks[0];

                    if (oJournalVouchers.Add() != 0)
                    {

                        strStatus = "E";

                        GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                        strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        return false;
                    }
                    else
                    {

                        strStatus = "S";

                        GlobalVariable.intErrNum = 0;
                        GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                        strPostDocNum = GlobalVariable.oCompany.GetNewObjectKey().ToString();
                        strMsgBod = string.Format("Successfully Posted {0} - {1}. Posted Document Number: {1} ", GlobalVariable.strDocType, GlobalVariable.strFileName, strPostDocNum);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, GlobalVariable.oCompany.GetNewObjectKey(), strPostDocNum, dteStart, "S", "0", strMsgBod);

                        if (GlobalVariable.oCompany.InTransaction)
                            GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                    }

                    GC.Collect();

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                return false;
            }
        }
        private static bool importDIAPIPostBP(string strEmpID, string strEmpName)
        {

            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.BusinessPartners oBusinessPartners;

            oRecordset = null;
            oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordset.DoQuery(string.Format("SELECT \"CardCode\" FROM OCRD WHERE \"U_EmployeeID\" = '{0}' ", strEmpID));

            if (!(oRecordset.RecordCount > 0))
            {
                oBusinessPartners = null;
                oBusinessPartners = (SAPbobsCOM.BusinessPartners)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                oBusinessPartners.Series = 61;
                oBusinessPartners.CardName = strEmpName;
                oBusinessPartners.CardType = BoCardTypes.cSupplier;
                oBusinessPartners.GroupCode = 105;

                oBusinessPartners.AccountRecivablePayables.SetCurrentLine(0);
                oBusinessPartners.AccountRecivablePayables.AccountCode = "20108-005-00";
                oBusinessPartners.AccountRecivablePayables.Add();

                if (oBusinessPartners.Add() != 0)
                {
                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                    return false;
                }
                else
                {
                    strCardCode = GlobalVariable.oCompany.GetNewObjectKey();
                }

            }

            return true;
            
        }
    }
}
