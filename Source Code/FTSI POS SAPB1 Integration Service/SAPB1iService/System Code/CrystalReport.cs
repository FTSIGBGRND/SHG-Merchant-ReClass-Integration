using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using SAPbobsCOM;

namespace SAPB1iService
{
    class CrystalReport
    {
        private static DateTime dteStart;
        private static string strTransType;
        public static bool CRBIR2307(string strFileName)
        {
            string strQuery;

            try
            {
                dteStart = DateTime.Now;

                strTransType = "CrystalReport Export -  BIR 2307";

                TableLogOnInfos CRTableLogoninfos = new TableLogOnInfos();
                TableLogOnInfo CRTableLogoninfo = new TableLogOnInfo();
                ConnectionInfo CRConnectionInfo = new ConnectionInfo();
                Tables CRTables;
                CrystalReportViewer CRViewer = new CrystalReportViewer();

                var cryRpt = new ReportDocument();          
                cryRpt.Load(GlobalVariable.strCRPath + "BIR 2307 - D.rpt");

                CRConnectionInfo.ServerName = GlobalVariable.strServer;
                CRConnectionInfo.DatabaseName = GlobalVariable.strSBOCompany;
                CRConnectionInfo.UserID = GlobalVariable.strDBUserName;
                CRConnectionInfo.Password = GlobalVariable.strDBPassword;

                cryRpt.SetParameterValue("USERCODE@", GlobalVariable.strSBOUserName);
                CRTables = cryRpt.Database.Tables;

                foreach (Table CRTable in CRTables)
                {
                    CRTableLogoninfo = CRTable.LogOnInfo;
                    CRTableLogoninfo.ConnectionInfo = CRConnectionInfo;
                    CRTable.ApplyLogOnInfo(CRTableLogoninfo);
                }

                CRViewer.ReportSource = cryRpt;

                ExportOptions CRExportOptions;
                DiskFileDestinationOptions CRDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CRFormatTypeOptions = new PdfRtfWordFormatOptions();

                CRDiskFileDestinationOptions.DiskFileName = strFileName;
                CRExportOptions = cryRpt.ExportOptions;

                CRExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CRExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                CRExportOptions.DestinationOptions = CRDiskFileDestinationOptions;
                CRExportOptions.FormatOptions = CRFormatTypeOptions;

                cryRpt.Export();

                strQuery = string.Format("DELETE FROM \"@FT_BIRPARAM\" WHERE \"U_UserCode\" = '{0}' AND \"U_RType\" = '2307' ", GlobalVariable.strSBOUserName);
                if (!(SystemFunction.executeQuery(strQuery)))
                {

                    GlobalVariable.intErrNum = -899;
                    GlobalVariable.strErrMsg = string.Format("Error updating BIR 2307 Parameters");

                    SystemFunction.transHandler("Crystal Report", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    return false;
                }

                cryRpt.Close();
                cryRpt.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Crystal Report", strTransType, "", "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }

        }
    }
}
