using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
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
    class WebAPIRequest
    {
        private static DateTime dteStart;
        public static bool postWebAPIRequest(string jsonData, string strUrl)
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    var url = new Uri(strUrl);

                    var payLoad = new StringContent(jsonData, Encoding.UTF8, "application/json");

                    var response = httpClient.PostAsync(url, payLoad).Result;

                    if (response.IsSuccessStatusCode)
                        return true;
                    else
                    {
                        GlobalVariable.intErrNum = -111;
                        GlobalVariable.strErrMsg = "Posting error occured while processing Post Web API Request";

                        return false;
                    }
                } 
            }
            catch (Exception ex)
            {
                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = "Exception error occured while processing Post Web API Request";

                return false;

            }
        }
        public static bool getWebAPICredentials()
        {
            string[] strLines;

            dteStart = DateTime.Now;

            try
            {
                strLines = File.ReadAllLines(GlobalVariable.strAPISettings);

                GlobalVariable.strBaseUrl = strLines[0].ToString().Substring(strLines[0].IndexOf("=") + 1);
                GlobalVariable.strAPIKey = strLines[1].ToString().Substring(strLines[1].IndexOf("=") + 1);

                if (string.IsNullOrEmpty(GlobalVariable.strBaseUrl))
                {
                    SystemFunction.transHandler("System", "API Settings", "", "", "", "", dteStart, "E", "-111", "Please check API Connection Settings.");
                    return false;
                }

            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("System", "API Settings", "", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
                return false;
            }

            return true;
        }
    }
}
