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
using SAPB1iService;

namespace SAPB1iService
{
    class Import
    {
        public static void _Import()
        {

            importFTPFiles();
     
            ImportDocuments._ImportDocuments();

            ImportPayments._ImportPayments();

        }
        private static void importFTPFiles()
        {
            string[] strImpExt;

            if (!string.IsNullOrEmpty(GlobalVariable.strImpExt))
            {
                strImpExt = GlobalVariable.strImpExt.Split(Convert.ToChar("|"));

                for (int intStr = 0; intStr < strImpExt.Length; intStr++)
                {
                    TransferFile.importSFTPFiles(strImpExt[intStr]);
                }
            }
        }
    }
}
