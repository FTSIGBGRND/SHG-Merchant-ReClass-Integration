using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using SAPB1iService;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data.SqlClient;

namespace SAPB1iService
{
    class ImportJournalEntry
    {
        private static DateTime dteStart = DateTime.Now;
        private static DataTable oDataTable;
        public static void _ImportJournalEntry(string strCompany)
        {
            importSQLJournalEntry(strCompany);
        }
        private static void importSQLJournalEntry(string strCompany)
        {
            SqlDataAdapter SqlDtaAdptr;

            try
            {
                dteStart = DateTime.Now;

                oDataTable = new DataTable();
                oDataTable.Clear();

                if (GlobalVariable.SapCon.State == ConnectionState.Closed)
                    GlobalVariable.SapCon.Open();

                SqlDtaAdptr = new SqlDataAdapter("EXEC FTSI_B1IS_IMPORT_FTFCS_JOURNALENTRY", GlobalVariable.SapCon);
                SqlDtaAdptr.Fill(oDataTable);

                if (GlobalVariable.SapCon.State == ConnectionState.Open)
                    GlobalVariable.SapCon.Close();

                importDIAPIPostJournalEntry();

            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", "Journal Entry", "30", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static void importDIAPIPostJournalEntry()
        {

            string strStatus = "", strMsgBod, strPostDocNum = "";

            string strQuery, strJERefCode = "", strVendorType, strTransType;

            string strARClrAcct = "", strSlsDscAcct = "", strMrcInvAcct = "",
                   strNetSlsAcct = "", strOutVatAcct = "", strExpClrAcct = "";

            double dblARClr, dblSlsDsc, dblMrcInv, dblNetSls, dblOutVat, dblExpClr;

            SAPbobsCOM.JournalEntries oJournalEntries;
            SAPbobsCOM.Recordset oRecordset;

            try
            {
                GlobalFunction.getObjType(30);

                if (oDataTable.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow <= oDataTable.Rows.Count - 1; intRow++)
                    {
                        strJERefCode = oDataTable.Rows[intRow]["JERefCode"].ToString();
                        strVendorType = oDataTable.Rows[intRow]["U_VendorType"].ToString();
                        strTransType = oDataTable.Rows[intRow]["U_TransType"].ToString();

                        dblARClr = Convert.ToDouble(oDataTable.Rows[intRow]["U_Sales"].ToString());
                        dblSlsDsc = Convert.ToDouble(oDataTable.Rows[intRow]["U_NVatDiscount"].ToString());
                        dblMrcInv = Convert.ToDouble(oDataTable.Rows[intRow]["U_Cost"].ToString());
                        dblNetSls = Convert.ToDouble(oDataTable.Rows[intRow]["U_NVatSales"].ToString());
                        dblOutVat = Convert.ToDouble(oDataTable.Rows[intRow]["U_OutputVat"].ToString());
                        dblExpClr = Convert.ToDouble(oDataTable.Rows[intRow]["ExpClr"].ToString());

                        strQuery = string.Format("SELECT \"U_ARClrAcct\", \"U_SlsDscAcct\", \"U_MrcInvAcct\", \"U_NetSlsAcct\", \"U_OutVatAcct\", \"U_ExpClrAcct\" " +
                                                 "FROM \"@FTFCSGL\" " + 
                                                 "WHERE \"Code\" = '{0}' ", strVendorType);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            strARClrAcct = oRecordset.Fields.Item("U_ARClrAcct").Value.ToString();
                            strSlsDscAcct = oRecordset.Fields.Item("U_SlsDscAcct").Value.ToString();
                            strMrcInvAcct = oRecordset.Fields.Item("U_MrcInvAcct").Value.ToString();
                            strNetSlsAcct = oRecordset.Fields.Item("U_NetSlsAcct").Value.ToString();
                            strOutVatAcct = oRecordset.Fields.Item("U_OutVatAcct").Value.ToString();
                            strExpClrAcct = oRecordset.Fields.Item("U_ExpClrAcct").Value.ToString();

                            oJournalEntries = null;
                            oJournalEntries = (SAPbobsCOM.JournalEntries)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

                            oJournalEntries.ReferenceDate = Convert.ToDateTime(oDataTable.Rows[intRow]["U_TransDate"].ToString());
                            oJournalEntries.DueDate = Convert.ToDateTime(oDataTable.Rows[intRow]["U_TransDate"].ToString());
                            oJournalEntries.TaxDate = Convert.ToDateTime(oDataTable.Rows[intRow]["U_TransDate"].ToString());
                            oJournalEntries.Memo = oDataTable.Rows[intRow]["U_Comments"].ToString();

                            oJournalEntries.UserFields.Fields.Item("U_BaseType").Value = oDataTable.Rows[intRow]["U_BaseType"].ToString();
                            oJournalEntries.UserFields.Fields.Item("U_BaseKey").Value = strJERefCode;
                            oJournalEntries.UserFields.Fields.Item("U_RegNum").Value = oDataTable.Rows[intRow]["U_RegNum"].ToString();
                            oJournalEntries.UserFields.Fields.Item("U_RefNum1").Value = oDataTable.Rows[intRow]["U_TransNum"].ToString();
                            oJournalEntries.UserFields.Fields.Item("U_RefNum2").Value = oDataTable.Rows[intRow]["U_RefNo"].ToString();
                            oJournalEntries.UserFields.Fields.Item("U_Vendor").Value = oDataTable.Rows[intRow]["U_Vendor"].ToString();


                            #region "AR CLEARING"

                            oJournalEntries.Lines.AccountCode = strARClrAcct;
                            if (strTransType == "S")
                            {
                                oJournalEntries.Lines.Debit = dblARClr;
                                oJournalEntries.Lines.Credit = 0;
                            }
                            else
                            {
                                oJournalEntries.Lines.Debit = 0;
                                oJournalEntries.Lines.Credit = dblARClr;
                            }

                            oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                            oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();
                            oJournalEntries.Lines.UserFields.Fields.Item("U_Desc").Value = oDataTable.Rows[intRow]["U_Comments"].ToString();

                            #endregion

                            if (strVendorType == "LC" || strVendorType == "IC" || strVendorType == "C")
                            {
                                #region "DISCOUNT"

                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strSlsDscAcct;
                                if (strTransType == "S")
                                {
                                    oJournalEntries.Lines.Debit = dblSlsDsc;
                                    oJournalEntries.Lines.Credit = 0;
                                }
                                else
                                {
                                    oJournalEntries.Lines.Debit = 0;
                                    oJournalEntries.Lines.Credit = dblSlsDsc;
                                }

                                oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                                oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();

                                #endregion

                                #region "SALES"

                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strNetSlsAcct;
                                if (strTransType == "S")
                                {
                                    oJournalEntries.Lines.Debit = 0;
                                    oJournalEntries.Lines.Credit = dblNetSls;
                                }
                                else
                                {
                                    oJournalEntries.Lines.Debit = dblNetSls;
                                    oJournalEntries.Lines.Credit = 0;
                                }

                                oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                                oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();

                                #endregion

                                #region "OUTPUT VAT"

                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strOutVatAcct;
                                if (strTransType == "S")
                                {
                                    oJournalEntries.Lines.Debit = 0;
                                    oJournalEntries.Lines.Credit = dblOutVat;
                                }
                                else
                                {
                                    oJournalEntries.Lines.Debit = dblOutVat;
                                    oJournalEntries.Lines.Credit = 0;
                                }

                                oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                                oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();

                                #endregion

                            }
                            else
                            {
                                #region "MERCHANDISE INVENTORY"

                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strMrcInvAcct;
                                if (strTransType == "S")
                                {
                                    oJournalEntries.Lines.Debit = 0;
                                    oJournalEntries.Lines.Credit = dblMrcInv;
                                }
                                else
                                {
                                    oJournalEntries.Lines.Debit = dblMrcInv;
                                    oJournalEntries.Lines.Credit = 0;
                                }

                                oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                                oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();

                                #endregion

                                #region "EXPENSE CLEARING"

                                oJournalEntries.Lines.Add();
                                oJournalEntries.Lines.AccountCode = strExpClrAcct;

                                if (strTransType == "S")
                                {
                                    oJournalEntries.Lines.Debit = 0;
                                    oJournalEntries.Lines.Credit = dblExpClr;
                                }
                                else
                                {
                                    oJournalEntries.Lines.Debit = dblExpClr;
                                    oJournalEntries.Lines.Credit = 0;
                                }

                                oJournalEntries.Lines.CostingCode2 = oDataTable.Rows[intRow]["LocCode"].ToString();
                                oJournalEntries.Lines.CostingCode4 = oDataTable.Rows[intRow]["DepCode"].ToString();

                                #endregion

                            }

                            if (oJournalEntries.Add() != 0)
                            {
                                GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                strStatus = "E";
                                strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, strJERefCode, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                SystemFunction.transHandler("Import", GlobalVariable.strDocType, GlobalVariable.intObjType.ToString(), strJERefCode, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);
                            }
                            else
                            {
                                strPostDocNum = GlobalFunction.getJENum(GlobalVariable.oCompany.GetNewObjectKey().ToString());

                                strStatus = "S";
                                strMsgBod = string.Format("Successfully Posted {0} - {1}. Posted Document Number: {2} ", GlobalVariable.strDocType, strJERefCode, strPostDocNum);

                                SystemFunction.transHandler("Import", GlobalVariable.strDocType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, GlobalVariable.oCompany.GetNewObjectKey(), strPostDocNum, dteStart, "S", "0", strMsgBod);

                                if (GlobalVariable.oCompany.InTransaction)
                                    GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                            }

                        }
                        else
                        {
                            GlobalVariable.intErrNum = -999;
                            GlobalVariable.strErrMsg = "G/L Account Setup not Found for FCS Journal Entry Integration.";

                            strStatus = "E";
                            strMsgBod = string.Format("Error Posting {0} - {1}.\rError Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, strJERefCode, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", GlobalVariable.strDocType, GlobalVariable.intObjType.ToString(), strJERefCode, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        }

                        updateBaseDoc(strStatus, strMsgBod, strJERefCode);

                        GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());
                    }

                }
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", GlobalVariable.strDocType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", "-111", ex.Message.ToString());

                if (!(string.IsNullOrEmpty(strJERefCode)))
                    updateBaseDoc("E", string.Format("Error Adding Posting Journal Entry - {0} ", ex.Message.ToString()), strJERefCode);
            }
        }
        private static bool updateBaseDoc(string strStatus, string strRemarks, string strDocNum)
        {
            SqlCommand SqlCom;

            string strQuery;

            try
            {

                if (GlobalVariable.SapCon.State == ConnectionState.Closed)
                    GlobalVariable.SapCon.Open();

                strQuery = string.Format("UPDATE \"@FTFCSJE\" SET \"U_Status\" = '{0}', \"U_Remarks\" = '{1}' WHERE Code = '{2}' ", strStatus, strRemarks.Replace("'", ""), strDocNum);
                SqlCom = new SqlCommand(strQuery, GlobalVariable.SapCon);
                SqlCom.ExecuteNonQuery();

                SqlCom.Dispose();

                if (GlobalVariable.SapCon.State == ConnectionState.Open)
                    GlobalVariable.SapCon.Close();

                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {
                SystemFunction.errorAppend(ex.Message.ToString());
                return false;
            }
        }

    }
}
