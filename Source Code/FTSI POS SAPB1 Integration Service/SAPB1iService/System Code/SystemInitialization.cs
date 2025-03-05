using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPB1iService;
using System.IO;

namespace SAPB1iService
{
    class SystemInitialization
    {
        public static bool initTables()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            if (SystemFunction.createUDT("FTPOSISL", "FT IEMOP Integration Log", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (SystemFunction.createUDT("FTISSP", "FT Integration SetUp", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            /****************************** UNTIL HERE - THANK YOU ***************************************************/


            return true;
        }
        public static bool initFields()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            #region "FRAMEWORK UDF"

            /******************************* INTEGRATION SERVICE LOG ***********************************************/

            if (SystemFunction.isUDFexists("@FTPOSISL", "Process") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "Process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "TransType") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "ObjType") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "TransDate") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "TransDate", "Transaction Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "FileName") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "TrgtDocKey") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "TrgtDocKey", "Base Document Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "TrgtDocNum") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "TrgtDocNum", "Base Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "StartTime") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "StartTime", "StartTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "EndTime") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "EndTime", "EndTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "Status") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "ErrorCode") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "ErrorCode", "Error Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPOSISL", "Remarks") == false)
                if (SystemFunction.createUDF("@FTPOSISL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            /******************************* INTEGRATION SETUP ***********************************************/

            if (SystemFunction.isUDFexists("@FTISSP", "ExportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportFile", "Export File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ExportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportPath", "Export Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportFile", "Import File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportPath", "Import Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "Delimiter") == false)
                if (SystemFunction.createUDF("@FTISSP", "Delimiter", "Delimiter", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcessTime") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcessTime", "Process Time", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "AlwaysRun") == false)
                if (SystemFunction.createUDF("@FTISSP", "AlwaysRun", "Services Always Running?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcSer") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcSer", "Process Service", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RunRep") == false)
                if (SystemFunction.createUDF("@FTISSP", "RunRep", "Reprocess Error File?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RepDate") == false)
                if (SystemFunction.createUDF("@FTISSP", "RepDate", "Reprocess Error Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

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

            #region "AUTOMATIC AR INVOICE"

            if (SystemFunction.isUDFexists("OCRC", "CardCode") == false)
                if (SystemFunction.createUDF("OCRC", "CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "CardName") == false)
                if (SystemFunction.createUDF("OCRC", "CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "RCVatGroup") == false)
                if (SystemFunction.createUDF("OCRC", "RCVatGroup", "ReClass Vat Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "BCGLAcct") == false)
                if (SystemFunction.createUDF("OCRC", "BCGLAcct", "Bank Charge GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "BCRate") == false)
                if (SystemFunction.createUDF("OCRC", "BCRate", "Bank Charge Rate", SAPbobsCOM.BoFldSubTypes.st_Rate, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "BCVatGroup") == false)
                if (SystemFunction.createUDF("OCRC", "BCVatGroup", "Bank Charge Vat Group", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OCRC", "WTCode") == false)
                if (SystemFunction.createUDF("OCRC", "WTCode", "WTax Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "ORCTDocNum") == false)
                if (SystemFunction.createUDF("OINV", "ORCTDocNum", "Incoming Payment DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "ORCTDocEnt") == false)
                if (SystemFunction.createUDF("OINV", "ORCTDocEnt", "Incoming Payment Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OADM", "InvSeries") == false)
                if (SystemFunction.createUDF("OADM", "InvSeries", "AR Invoice Series", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            #endregion

            /************************** ITEM MASTER DATA ***************************************************************/

            if (SystemFunction.isUDFexists("OITM", "isExtract") == false)
                if (SystemFunction.createUDF("OITM", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;

            /************************** BUSINESS PARTNER DATA **********************************************************/

            if (SystemFunction.isUDFexists("OCRD", "isExtract") == false)
                if (SystemFunction.createUDF("OCRD", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;


            /************************** INCOMING PAYMENT **********************************************************/

            if (SystemFunction.isUDFexists("ORCT", "isExtract") == false)
                if (SystemFunction.createUDF("ORCT", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, S - Success", "") == false)
                    return false;

            /************************** ADMINISTRATION ****************************************************************/

            if (SystemFunction.isUDFexists("OUSR", "IntMsg") == false)
                if (SystemFunction.createUDF("OUSR", "IntMsg", "Integration Message", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OADM", "Company") == false)
                if (SystemFunction.createUDF("OADM", "Company", "Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            #endregion

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

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
            catch(Exception ex)
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
        public static bool initSQLConnection()
        {
            if (File.Exists(GlobalVariable.strSQLSettings))
            {
                if (SystemFunction.connectSQL(GlobalVariable.strSQLSettings))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

    }
}
