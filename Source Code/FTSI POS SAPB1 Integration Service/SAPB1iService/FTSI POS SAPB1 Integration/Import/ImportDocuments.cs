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

namespace SAPB1iService
{
    class ImportDocuments
    {
        private static DateTime dteStart;
        private static string strTransType;

        private static string strObjType, strVersion, strRefNum;
        private static string strMsgBod, strStatus, strPostDocNum, strPostDocEnt;

        private static bool blExist = false;
        public static void _ImportDocuments()
        {
            
            importFromXMLFiles();
            
            ImportUserDefinedDocuments._ImportUserDefinedDocuments();

        }
        private static void importFromXMLFiles()
        {
            string[] strFValue;

            try
            {

                dteStart = DateTime.Now;

                strTransType = "Documents - Import XML";

                string[] strFileImport = new string[] { string.Format("*{0}_{1}_*.xml", GlobalVariable.strCompany, "DOC"), string.Format("*{0}_{1}_*.XML", GlobalVariable.strCompany, "DOC") };

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, fileimport))
                    {
                        dteStart = DateTime.Now;

                        GlobalVariable.strFileName = Path.GetFileName(strFile);

                        strFValue = Path.GetFileNameWithoutExtension(strFile).Split(Convert.ToChar("_"));

                        strRefNum = strFValue[4];
                        strVersion = strFValue[6];

                        importXMLPostDocument(strFile);
                    }
                }

                GC.Collect();
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

            }
        }
        private static void importXMLPostDocument(string strFile)
        {

            SAPbobsCOM.Documents oDocuments;

            XmlDocument xmlDoc = new XmlDocument();

            try
            { 
                if (validateXMLData(strFile, strVersion))
                {
                    if (!(GlobalVariable.oCompany.InTransaction))
                        GlobalVariable.oCompany.StartTransaction();

                    oDocuments = null;
                    oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObjectFromXML(strFile, 0);

                    if (blExist == false)
                    {
                        if (oDocuments.Add() != 0)
                        {
                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                            strStatus = "E";
                            strMsgBod = string.Format("Error Posting {0} - {1}.\r" +
                                                      "Error Code: {2}\rDescription: {3} ", GlobalVariable.strDocType, GlobalVariable.strFileName, 
                                                                                            GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        }
                        else
                        {
                            strStatus = "S";

                            strPostDocEnt = GlobalVariable.oCompany.GetNewObjectKey().ToString();
                            strPostDocNum = GlobalFunction.getDocNum(GlobalVariable.intObjType, strPostDocEnt);

                            strMsgBod = string.Format("Successfully Posted {0} - {1}. Posted Document Number: {1} ", GlobalVariable.strDocType, GlobalVariable.strFileName, strPostDocNum);

                            GlobalVariable.intErrNum = 0;
                            GlobalVariable.strErrMsg = strMsgBod;

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, strPostDocEnt, strPostDocNum, dteStart, strStatus, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            if (GlobalVariable.oCompany.InTransaction)
                                GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                        }

                        TransferFile.transferProcFiles("Import", strStatus, GlobalVariable.strFileName);

                        GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                    }
                    //else condition here if there is an update with the documents (need to define mapping because xml update is not working specially with partial transactions made
                }
                else
                {               
                    TransferFile.transferProcFiles("Import", "E", GlobalVariable.strFileName);
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                TransferFile.transferProcFiles("Import", "E", GlobalVariable.strFileName);
            }
}
        private static bool validateXMLData(string strFilePath, string strVersion)
        {
            string strQuery;

            string strBaseLine = "", strBaseRef = "", strBaseType = "", strB1BaseLine, strB1BaseEntry,
                   strCodeBars;

            bool blRetVal = true, blSaveDoc = false;

            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xmlNodLst;

            SAPbobsCOM.Recordset oRecordset;

            try
            {

                xmlDoc.Load(strFilePath);

                xmlNodLst = xmlDoc.SelectNodes("BOM/BO/AdmInfo");
                foreach (XmlNode xmlNod in xmlNodLst)
                {
                    strObjType = xmlNod.SelectSingleNode("Object").InnerText;
                    GlobalFunction.getObjType(Convert.ToInt32(strObjType));
                }

                strQuery = string.Format("SELECT \"DocEntry\" FROM {0} WHERE \"U_FileName\" = '{1}' ", GlobalVariable.strTableHeader, GlobalVariable.strFileName);

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                {
                    GlobalVariable.strErrMsg = string.Format("{0} \rFile Already Uploaded - {1}.", GlobalVariable.strErrMsg, GlobalVariable.strFileName);
                    blRetVal = false;
                }

                strQuery = string.Format("SELECT \"DocEntry\" FROM {0} WHERE \"U_RefNum\" = '{1}' ", GlobalVariable.strTableHeader, strRefNum);

                oRecordset = null;
                oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRecordset.DoQuery(strQuery);

                if (oRecordset.RecordCount > 0)
                    blExist = true;

                xmlNodLst = xmlDoc.SelectNodes(string.Format("BOM/BO/{0}/row", GlobalVariable.strTableHeader));
                foreach (XmlNode xmlNod in xmlNodLst)
                {
                    //validation header if needed
                }

                xmlNodLst = xmlDoc.SelectNodes(string.Format("BOM/BO/{0}/row", GlobalVariable.strTableLine1));
                foreach (XmlNode xmlNod1 in xmlNodLst)
                {
                    //validation details if needed

                    if (xmlNod1.SelectSingleNode("U_BaseType") != null)
                        strBaseType = xmlNod1.SelectSingleNode("U_BaseType").InnerText;

                    if (xmlNod1.SelectSingleNode("U_BaseRef") != null)
                        strBaseRef = xmlNod1.SelectSingleNode("U_BaseRef").InnerText;

                    if (xmlNod1.SelectSingleNode("U_BaseLine") != null)
                        strBaseLine = xmlNod1.SelectSingleNode("U_BaseLine").InnerText;

                    if (!(string.IsNullOrEmpty(strBaseRef)))
                    {
                        #region "Remove Existing Base Reference / Unnecessary Fields"

                        if (xmlNod1.SelectSingleNode("CodeBars") != null)
                        {
                            strCodeBars = xmlNod1.SelectSingleNode("CodeBars").InnerText;

                            if (string.IsNullOrEmpty(strCodeBars))
                            {
                                XmlNode CodeBars = xmlNod1.SelectSingleNode("CodeBars");
                                xmlNod1.RemoveChild(CodeBars);
                                blSaveDoc = true;
                            }
                        }

                        if (xmlNod1.SelectSingleNode("BaseLine") != null)
                        {
                            XmlNode BaseLine = xmlNod1.SelectSingleNode("BaseLine");
                            xmlNod1.RemoveChild(BaseLine);
                            blSaveDoc = true;
                        }

                        if (xmlNod1.SelectSingleNode("BaseRef") != null)
                        {
                            XmlNode BaseRef = xmlNod1.SelectSingleNode("BaseRef");
                            xmlNod1.RemoveChild(BaseRef);
                            blSaveDoc = true;
                        }

                        if (xmlNod1.SelectSingleNode("BaseEntry") != null)
                        {
                            XmlNode BaseEntry = xmlNod1.SelectSingleNode("BaseEntry");
                            xmlNod1.RemoveChild(BaseEntry);
                            blSaveDoc = true;
                        }

                        if (xmlNod1.SelectSingleNode("BaseType") != null)
                        {
                            XmlNode BaseType = xmlNod1.SelectSingleNode("BaseType");
                            xmlNod1.RemoveChild(BaseType);
                            blSaveDoc = true;
                        }

                        #endregion

                        GlobalFunction.getBaseType(Convert.ToInt32(strBaseType));

                        strQuery = string.Format("SELECT {0}.\"DocEntry\", {0}.\"DocNum\", {1}.\"LineNum\" " +
                                                 "FROM {0} INNER JOIN {1} ON {0}.\"DocEntry\" = {1}.\"DocEntry\" " +
                                                 "WHERE {0}.\"U_RefNum\" = '{2}' AND " +
                                                 "      {1}.\"LineNum\" = '{3}' ", GlobalVariable.strBTableHeader, GlobalVariable.strBTableLine1, strBaseRef, strBaseLine);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            XmlNode NBaseType = xmlDoc.CreateElement("BaseType");
                            NBaseType.InnerText = GlobalVariable.intBObjType.ToString();
                            xmlNod1.PrependChild(NBaseType);
                            blSaveDoc = true;

                            XmlNode NBaseLine = xmlDoc.CreateElement("BaseLine");
                            NBaseLine.InnerText = oRecordset.Fields.Item("LineNum").Value.ToString();
                            xmlNod1.PrependChild(NBaseLine);
                            blSaveDoc = true;

                            XmlNode NBaseRef = xmlDoc.CreateElement("BaseRef");
                            NBaseRef.InnerText = oRecordset.Fields.Item("DocNum").Value.ToString();
                            xmlNod1.PrependChild(NBaseRef);
                            blSaveDoc = true;

                            XmlNode NBaseEntry = xmlDoc.CreateElement("BaseEntry");
                            NBaseEntry.InnerText = oRecordset.Fields.Item("DocEntry").Value.ToString();
                            xmlNod1.PrependChild(NBaseEntry);
                            blSaveDoc = true;

                        }
                        else
                        {
                            GlobalVariable.strErrMsg = string.Format("{0} \rBase Document Not Found for Base Reference {0} - Line # {1} with Filename {2}.", GlobalVariable.strErrMsg, strBaseRef, strBaseLine, GlobalVariable.strFileName);

                            blSaveDoc = false;
                            blRetVal = false;
                        }

                    }
                }

                xmlNodLst = xmlDoc.SelectNodes(string.Format("BOM/BO/{0}/row", GlobalVariable.strTableLine3));
                foreach (XmlNode xmlNod3 in xmlNodLst)
                {
                    //validation details if needed

                    if (xmlNod3.SelectSingleNode("U_BaseType") != null)
                        strBaseType = xmlNod3.SelectSingleNode("U_BaseType").InnerText;

                    if (xmlNod3.SelectSingleNode("U_BaseRef") != null)
                        strBaseRef = xmlNod3.SelectSingleNode("U_BaseRef").InnerText;

                    if (xmlNod3.SelectSingleNode("U_BaseLine") != null)
                        strBaseLine = xmlNod3.SelectSingleNode("U_BaseLine").InnerText;

                    if (!(string.IsNullOrEmpty(strBaseRef)))
                    {
                        #region "Remove Existing Base Reference / Unnecessary Fields"

                        if (xmlNod3.SelectSingleNode("BaseLine") != null)
                        {
                            XmlNode BaseLine = xmlNod3.SelectSingleNode("BaseLine");
                            xmlNod3.RemoveChild(BaseLine);
                            blSaveDoc = true;
                        }

                        if (xmlNod3.SelectSingleNode("BaseRef") != null)
                        {
                            XmlNode BaseRef = xmlNod3.SelectSingleNode("BaseRef");
                            xmlNod3.RemoveChild(BaseRef);
                            blSaveDoc = true;
                        }

                        if (xmlNod3.SelectSingleNode("BaseEntry") != null)
                        {
                            XmlNode BaseEntry = xmlNod3.SelectSingleNode("BaseEntry");
                            xmlNod3.RemoveChild(BaseEntry);
                            blSaveDoc = true;
                        }

                        if (xmlNod3.SelectSingleNode("BaseType") != null)
                        {
                            XmlNode BaseType = xmlNod3.SelectSingleNode("BaseType");
                            xmlNod3.RemoveChild(BaseType);
                            blSaveDoc = true;
                        }

                        #endregion

                        GlobalFunction.getBaseType(Convert.ToInt32(strBaseType));

                        strQuery = string.Format("SELECT {0}.\"DocEntry\", {0}.\"DocNum\", {1}.\"LineNum\" " +
                                                 "FROM {0} INNER JOIN {1} ON {0}.\"DocEntry\" = {1}.\"DocEntry\" " +
                                                 "WHERE {0}.\"U_RefNum\" = '{2}' AND " +
                                                 "      {1}.\"LineNum\" = '{3}' ", GlobalVariable.strBTableHeader, GlobalVariable.strBTableLine3, strBaseRef, strBaseLine);

                        oRecordset = null;
                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oRecordset.DoQuery(strQuery);

                        if (oRecordset.RecordCount > 0)
                        {
                            XmlNode NBaseType = xmlDoc.CreateElement("BaseType");
                            NBaseType.InnerText = GlobalVariable.intBObjType.ToString();
                            xmlNod3.PrependChild(NBaseType);
                            blSaveDoc = true;

                            XmlNode NBaseLine = xmlDoc.CreateElement("BaseLnNum");
                            NBaseLine.InnerText = oRecordset.Fields.Item("LineNum").Value.ToString();
                            xmlNod3.PrependChild(NBaseLine);
                            blSaveDoc = true;

                            XmlNode NBaseRef = xmlDoc.CreateElement("BaseRef");
                            NBaseRef.InnerText = oRecordset.Fields.Item("DocNum").Value.ToString();
                            xmlNod3.PrependChild(NBaseRef);
                            blSaveDoc = true;

                            XmlNode NBaseEntry = xmlDoc.CreateElement("BaseAbsEnt");
                            NBaseEntry.InnerText = oRecordset.Fields.Item("DocEntry").Value.ToString();
                            xmlNod3.PrependChild(NBaseEntry);
                            blSaveDoc = true;

                        }
                        else
                        {
                            GlobalVariable.strErrMsg = string.Format("{0} \rBase Document Not Found for Base Reference {0} - Line # {1} with Filename {2}.", GlobalVariable.strErrMsg, strBaseRef, strBaseLine, GlobalVariable.strFileName);

                            blSaveDoc = false;
                            blRetVal = false;
                        }

                    }
                }

                GC.Collect();

                if (blSaveDoc == true)
                    xmlDoc.Save(strFilePath);

                if (blRetVal == false)
                {
                    SystemFunction.transHandler("Import", "Documents", strObjType, GlobalVariable.strFileName, "", "", dteStart, "E", "-999", GlobalVariable.strErrMsg);
                    return false;
                }
                else
                    return true;

            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", "Documents", GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }
        }
    }
}
