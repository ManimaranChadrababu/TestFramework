using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NewFrameWork.Selenium
{
    public static class ExtendedDriver
    {
        static int i = 0;

        public static void Quits(this IWebDriver driver)
        {
            ICapabilities capabilities = ((RemoteWebDriver)driver).Capabilities;
            string browser = capabilities.BrowserName;
            if (browser.Equals("chrome", System.StringComparison.InvariantCultureIgnoreCase))
            {
                driver.Quit();
            }

            else if (browser.Equals("ie", System.StringComparison.InvariantCultureIgnoreCase))
            {
                driver.Quit();
            }

            else
            {
                try
                {
                    for (int i = 0; i <= 2; i++)
                    {

                        Process[] proc = Process.GetProcesses();
                        Process[] proc2 = Process.GetProcessesByName("WerFault");
                        if (proc2.Length > 0)
                        {
                            try
                            {
                                proc2[0].Kill();
                            }
                            catch (Exception e)
                            {
                            }
                            break;
                        }

                        else
                        {
                            Thread.Sleep(3000);
                            i++;
                        }

                    }

                    for (int i = 0; i <= 2; i++)
                    {

                        Process[] proc = Process.GetProcesses();
                        Process[] proc2 = Process.GetProcessesByName("csrss");
                        if (proc2.Length > 0)
                        {
                            foreach (Process s in proc2)
                            {
                                if (s.MainWindowTitle.Contains("plugin-container"))
                                {
                                    try
                                    {
                                        bool status = s.CloseMainWindow();
                                       
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }

                        else
                        {
                            Thread.Sleep(2000);
                            i++;
                        }

                    }

                    for (int i = 0; i <= 2; i++)
                    {

                        Process[] proc = Process.GetProcesses();
                        Process[] proc2 = Process.GetProcessesByName("plugin-container");
                        if (proc2.Length > 0)
                        {
                            foreach (Process s in proc2)
                            {
                                if (s.MainWindowTitle.Contains("plugin-container"))
                                {
                                    try
                                    {
                                        bool status = s.CloseMainWindow();
                                      
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }
                        }

                        else
                        {
                            Thread.Sleep(2000);
                            i++;
                        }

                    }

                    driver.Close();
                }
                catch (Exception ex)
                {
                }
                for (int i = 0; i <= 2; i++)
                {

                    Process[] proc = Process.GetProcesses();
                    Process[] proc2 = Process.GetProcessesByName("WerFault");
                    if (proc2.Length > 0)
                    {
                        try
                        {
                            proc2[0].Kill();
                        }
                        catch (Exception e)
                        {
                        }
                        break;
                    }

                    else
                    {
                        Thread.Sleep(5000);
                        i++;
                    }

                }

                for (int i = 0; i <= 2; i++)
                {

                    Process[] proc = Process.GetProcesses();
                    Process[] proc2 = Process.GetProcessesByName("csrss");
                    if (proc2.Length > 0)
                    {
                        foreach (Process s in proc2)
                        {
                            if (s.MainWindowTitle.Contains("plugin-container"))
                            {
                                try
                                {
                                    //bool status = s.CloseMainWindow();
                                    s.Kill();
                                    s.Close();
                                   
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
                    }

                    else
                    {
                        Thread.Sleep(2000);
                        i++;
                    }

                }

                for (int i = 0; i <= 2; i++)
                {

                    Process[] proc = Process.GetProcesses();
                    Process[] proc2 = Process.GetProcessesByName("plugin-container");
                    if (proc2.Length > 0)
                    {
                        foreach (Process s in proc2)
                        {
                            if (s.MainWindowTitle.Contains("plugin-container"))
                            {
                                try
                                {
                                    //bool status = s.CloseMainWindow();
                                    s.Kill();
                                    s.Close();
                                    
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
                    }

                    else
                    {
                        Thread.Sleep(2000);
                        i++;
                    }

                }

                try
                {
                    driver.Dispose();
                    Process[] proc = Process.GetProcesses();
                    Process[] proc2 = Process.GetProcessesByName("firefox");
                    if (proc2.Length > 0)
                    {
                        foreach (Process s in proc2)
                        {

                            try
                            {
                                //bool status = s.CloseMainWindow();
                                s.Kill();
                                s.Close();
                               
                            }
                            catch (Exception ex)
                            {

                            }

                        }
                    }

                }
                catch (Exception ex)
                {
                }
                
            }

        }
    }
}
