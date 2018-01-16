using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NewFrameWork.Selenium;

namespace NewFrameWork.Generic
{
    class HtmlReport
    {
        public HtmlReport(string fileName1, bool timestamp)
        {
            if (timestamp)
            {
                string datetimeString = string.Format("{0:yyyyMMdd}", DateTime.Now);
                FileName = fileName1 + "-" + datetimeString + ".html";
            }
            else
            {
                FileName = fileName1 + ".html";
            }

        }
        /// <summary>
        ///This method creates the table header for the report to be stored
        /// </summary>
        /// <param name="headerData">A String array with the required table headers</param>
        public void AddHeader(string[] headerData)
        {
            string reportContent2 = "";
            int colLength = headerData.Length;

            if (colLength > 0)
            {
                reportContent2 = reportContent2 + "<br><br> \n";
                //reportContent2 = reportContent2 + "<center>";
                reportContent2 = reportContent2 + "<table id ='results' class=" + "sortable" +
                                 " width:auto cellpadding=2 cellspacing=0 border=" + "1" + ">";
                reportContent2 = reportContent2 + "<tbody>";
                reportContent2 = reportContent2 + "<tr>";

                int i = 0;

                while (i < headerData.Length)
                {
                    if (i == 3 || i == 4)
                    {
                        reportContent2 = reportContent2 + "<td class=bborder_left width=30%>" + headerData[i] + "</td>";
                    }
                    else
                    {
                        reportContent2 = reportContent2 + "<td class=bborder_left width=7%>" + headerData[i] + "</td>";
                    }
                    i++;
                }
                reportContent2 = reportContent2 + "</tr>";
            }

            using (var sw = new StreamWriter(FileName, true))
            {
                sw.WriteLine(reportContent2);
            }
        }

        /// <summary>
        /// This method adds the rows with test reports data
        /// </summary>
        /// <param name="reportData">A String array with the required report data.</param>
        public void AddReportData(string[] reportData)
        {
            string reportContent2 = "";
            int lenRep = reportData.Length;

            while (lenRep < reportData.Length)
            {
                reportData[lenRep] = "";
                lenRep++;
            }

            int i = 0;
            while (i < reportData.Length)
            {
                if (reportData[i] == null)
                {
                    reportData[i] = "---";
                }
                else if (reportData[i].Equals(string.Empty))
                {
                    reportData[i] = "---";
                }
                i++;
            }

            if (reportData.Length > 0)
            {
                reportContent2 = reportContent2 + "<tr>";

                int j = 0;

                while (j < reportData.Length)
                {
                    if (reportData[j].Equals("Pass", StringComparison.CurrentCultureIgnoreCase))
                    {
                        reportContent2 = reportContent2 + "<td class=border_left bgcolor=green><p class=normal_text>" +
                                         reportData[j] + "</p></td>";
                    }
                    else
                    {
                        if (reportData[j].Equals("Fail", StringComparison.CurrentCultureIgnoreCase))
                        {
                            reportContent2 = reportContent2 + "<td class=border_left bgcolor=red><p class=normal_text>" +
                                             reportData[j] + "</p></td>";
                        }
                        else
                        {
                            reportContent2 = reportContent2 + "<td class=border_left><p class=normal_text>" +
                                             reportData[j] + "</p></td>";
                        }
                    }
                    j++;
                }
            }
            reportContent2 = reportContent2 + "</tr>";

            using (var sw = new StreamWriter(FileName, true))
            {
                sw.WriteLine(reportContent2);
            }
        }

        /// <summary>
        ///This method prints the methods listed in Diffsectionarraay
        /// </summary>
        public void AddDiffsectionReport(List<string> Diffsectionarray)
        {

            string text = "";
            if (Diffsectionarray.Count != 0)
            {
                if (Diffsectionarray.Count == 1)
                {
                    text = Diffsectionarray[0];
                }
                else
                {
                    foreach (string s in Diffsectionarray)
                    {
                        text = text + s + ", ";
                    }
                }
            }

            else
            {
                text = "Changes not required in script";
            }

            using (var sw = new StreamWriter(FileName, true))
            {
                sw.WriteLine("<p align=left><b>  List of methods to be removed</b></p>");
                sw.WriteLine("<p align=left>" + text + "</p>");

            }

        }

        /// <summary>
        /////This method prints the Data index used by the test case
        ///// </summary>
        //public void AddDataindexReport(string Datasetindex, string[,] Data)
        //{
        //    List<string> datadetails = ReferenceFile.ReadRow(Datasetindex, Data);

        //    string text = "";
        //    if (datadetails.Count != 0)
        //    {
        //        if (datadetails.Count == 1)
        //        {
        //            text = datadetails[0];
        //        }
        //        else
        //        {
        //            foreach (string s in datadetails)
        //            {
        //                text = text + s + ", ";
        //            }
        //        }
        //    }

        //    else
        //    {
        //        text = "Dataset details not available";
        //    }

        //    using (var sw = new StreamWriter(FileName, true))
        //    {
        //        sw.WriteLine("<p align=left><b>  Dataset details</b></p>");
        //        sw.WriteLine("<p align=left>" + text + "</p>");

        //    }

        //}


        /// <summary>
        ///This method closes the HTML report writing and prints the overall status
        /// </summary>
        public void CloseReport()
        {
            if (File.Exists(FileName))
            {
                using (var sw = new StreamWriter(FileName, true))
                {
                    sw.WriteLine("</body>");
                    sw.WriteLine("</html>");
                }
            }
        }

        /// <summary>
        /// This function calls the LogInit and the InitReport function
        /// </summary>
        public void InitHtmlReport(string strModule)
        {
            LogInit();
            InitReport(strModule);
        }

        /// <summary>
        ///This method creates the basic structure of the HTML file (CSS and headers)
        /// </summary>
        private void InitReport(string strModule)
        {
            if (File.Exists(FileName))
            {
                using (var sw = new StreamWriter(FileName, true))
                {
                    sw.WriteLine("<head>");
                    sw.WriteLine("      <meta content=text/html; charset=ISO-8859-1 http-equiv=content-type>");
                    sw.WriteLine("      <title>Ranorex Test Report</title>");
                    sw.WriteLine("      <style type=text/css>");
                    sw.WriteLine(
                        "       .title {  font-family: 'Lobster', Georgia, Times, serif; font-size: 40px;  font-weight: bold; color:#806D7E;}");
                    sw.WriteLine("          .bold_text {  font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px;  font-weight: bold;}");
                    sw.WriteLine("           ..normal_text {  font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px;  font-weight: normal;}");
                    sw.WriteLine("           .small_text {  font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 10px;  font-weight: normal; }");
                    sw.WriteLine("          .border { border: 1px solid #000000;}");
                    sw.WriteLine("        .border_left { border-top: 1px solid #000000; border-left: 1px solid #000000; border-right: 1px solid #000000;}");
                    sw.WriteLine("       .border_right { border-top: 1px solid #045AFD; border-right: 1px solid #000000;}");
                    sw.WriteLine("       .result_ok { font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px;  font-weight: bold; background-color:green;text-align: center; }");
                    sw.WriteLine("       .result_nok { font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px;  font-weight: bold;background-color:red; text-align: center; }");
                    sw.WriteLine("       .overall_ok { font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px; background-color:green; font-weight: bold; text-align: left; }");
                    sw.WriteLine("       .overall_nok { font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px; background-color:red; font-weight: bold; text-align: left; }");
                    sw.WriteLine("        .bborder_left { border-top: 1px solid #000000; border-left: 1px solid #000000; border-bottom: 1px solid #000000; background-color:#000000;font-family: Segoe UI; font-size: 12px;  font-weight: bold;text-align: center; color: white;}");
                    sw.WriteLine("       .bborder_right { border-right: 1px solid #045AFD; background-color:#000000;font-family: 'Adobe Caslon Pro', 'Hoefler Text', Georgia, Garamond, Times, serif; font-size: 12px;  font-weight: bold; text-align: center; color: white;}");
                    sw.WriteLine("       .bborder_result { border-top: 1px solid #000000; border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-color: 	#FFFFFF;font-family: Segoe UI; font-size: 12px;  font-weight: bold;text-align: center; color: black;}");
                    sw.WriteLine("      </style>");
                    sw.WriteLine("<script src=" + "http://www.kryogenix.org/code/browser/sorttable/sorttable.js" +
                                 " type=" + "text/javascript" + "></script>");
                    sw.WriteLine("<script type=" + "'text/javascript'" + ">");

                    sw.WriteLine("var rowVisible = true;");

                    sw.WriteLine("function toggleDisplay(tbl) {");
                    sw.WriteLine("   tbl.style.display=''" + ";");
                    sw.WriteLine("var tblRows = tbl.rows;");

                    sw.WriteLine("   for (i = 0; i < tblRows.length; i++) {");
                    sw.WriteLine("      if (tblRows[i].className != " + "'headerRow'" + ")" + " {");
                    sw.WriteLine(" tblRows[i].style.display = (rowVisible) ? 'none':'' ;");
                    sw.WriteLine("      }");
                    sw.WriteLine("   }");
                    sw.WriteLine("   rowVisible = !rowVisible;");
                    sw.WriteLine("}");
                    sw.WriteLine("</script>");
                    sw.WriteLine("      </head>");
                    sw.WriteLine("      <body>");
                    sw.WriteLine("      <br>");
                    sw.WriteLine("      <center>");
                    sw.WriteLine("      <table width=95% border=0 cellpadding=2 cellspacing=2>");
                    sw.WriteLine("      <tbody>");
                    sw.WriteLine("      <tr>");
                    sw.WriteLine("      <td>");
                    sw.WriteLine("      <table width=100% border=0 cellpadding=2 cellspacing=2>");
                    sw.WriteLine("      <tbody>");
                    sw.WriteLine("      <tr>");
                    sw.WriteLine(
                        "      <td align=center><p class=title>Ranorex Test Report</p></td></tr><tr> <td align=left></img><td align=right><img src=" +
                        "http://www.merge.com/MergeHealthcare/media/design/logo_merge_healthcare.gif" +
                        "></img></td></tr>");
                    sw.WriteLine("      </tbody>");
                    sw.WriteLine("      </table>");
                    sw.WriteLine("      <br><br>");

                    sw.WriteLine("       <table class=testRunDetails align='left'  width=40%> ");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Portal Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Portalserver"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + "1.0" + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("        <td class=bborder_left>LTA Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["LTA"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["LTAVersion"] + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Gateway1 Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Gateway1Server"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Gateway1Version"] + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Gateway2 Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Gateway2Server"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Gateway2Version"] + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Pacs Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["PacsServer"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["Pacsversion"] + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>ICA Server</td>");
                    sw.WriteLine("        <td class=bborder_result>" + ConfigurationManager.AppSettings["ICA"] + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" +"1.0" + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("         <td class=bborder_left>Host Name</td>");
                    sw.WriteLine("        <td class=bborder_result>" + System.Net.Dns.GetHostName() + "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + OSName() + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("        <td class=bborder_left >Browser Name</td>");
                    sw.WriteLine("        <td class=bborder_result>" + Driver.browserName+ "</td>");
                    sw.WriteLine("        <td class=bborder_result>" + Driver.browserVersion + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("       <tr>");
                    sw.WriteLine("        <td class=bborder_left >Automation tool</td>");
                    sw.WriteLine("        <td class=bborder_result>" + "Selenium" + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Execution Time</td>");
                    sw.WriteLine("        <td class=bborder_result>" + string.Format("{0:yyyyMMdd-hhmmss}", DateTime.Now) + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("         <tr>");
                    sw.WriteLine("         <td class=bborder_left>Module</td>");
                    sw.WriteLine("        <td class=bborder_result>" + strModule + "</td>");
                    sw.WriteLine("       </tr>");
                    sw.WriteLine("      </table> ");
                    sw.WriteLine("      <br><br><br><br> ");
                }
            }
        }

        /// <summary>
        /// This method checks if the fileName already exists ,if not creates the physical file.
        /// </summary>
        private void LogInit()
        {
            if (!File.Exists(FileName))
            {
                using (File.Create(FileName))
                {
                }
            }
            else
            {
                File.Delete(FileName);
                using (File.Create(FileName))
                {
                }
            }
        }

        private static string HKLM_GetString(string path, string key)
        {
            try
            {
                RegistryKey rk = Registry.LocalMachine.OpenSubKey(path);
                if (rk == null) return "";
                return (string)rk.GetValue(key);
            }
            catch { return ""; }
        }

        private static string OSName()
        {
            string ProductName = HKLM_GetString(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName");
            string CSDVersion = HKLM_GetString(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion");
            if (ProductName != "")
            {
                return (ProductName.StartsWith("Microsoft") ? "" : "Microsoft ") + ProductName +
                            (CSDVersion != "" ? " " + CSDVersion : "");
            }
            return "";
        }

        public string FileName { get; set; }
    }
}
