using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace NewFrameWork.Selenium
{
   public class Driver
    {
        private static IWebDriver driver;
        public static string browserName;
        public static string browserVersion;

        public static IWebDriver Instance
        {
            get
            {

                return driver ?? (driver = NewInstance());
            }
            set { driver = value; }
        }

        public static IWebDriver NewInstance()
        {
            string browser = ConfigurationManager.AppSettings["browser"];
            if (browser.Equals("chrome", System.StringComparison.InvariantCultureIgnoreCase))
            {
                ChromeOptions options = new ChromeOptions();
                options.Proxy = null;
                //options.AddArgument("Proxy = null");
                options.AddArgument("--start-maximized");
                options.AddArguments("--always-authorize-plugins=true");
                driver = new ChromeDriver(ConfigurationManager.AppSettings["drivers"], options);
                browserName = ((RemoteWebDriver)driver).Capabilities.BrowserName;
                browserVersion = ((RemoteWebDriver)driver).Capabilities.Version;
            }
            else if (browser.Equals("ie", System.StringComparison.InvariantCultureIgnoreCase))
            {
                InternetExplorerOptions IEInst = new InternetExplorerOptions();
                driver = new InternetExplorerDriver(ConfigurationManager.AppSettings["drivers"], IEInst);
                driver.Manage().Window.Maximize();
                browserName = ((RemoteWebDriver)driver).Capabilities.BrowserName;
                browserVersion = ((RemoteWebDriver)driver).Capabilities.Version;
            }

            else
            {

                string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
                string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default*", SearchOption.TopDirectoryOnly);
                if (pathsToProfiles.Length != 0)
                {
                    FirefoxProfile profile = new FirefoxProfile(pathsToProfiles[0]);
                    // profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                    driver = new FirefoxDriver(new FirefoxBinary(), profile);
                    driver.Manage().Window.Maximize();
                    browserName = ((RemoteWebDriver)driver).Capabilities.BrowserName;
                    browserVersion = ((RemoteWebDriver)driver).Capabilities.Version;
                }
                else
                {
                    driver = new FirefoxDriver();
                    driver.Manage().Window.Maximize();
                    browserName = ((RemoteWebDriver)driver).Capabilities.BrowserName;
                    browserVersion = ((RemoteWebDriver)driver).Capabilities.Version;
                }

            }

            return driver;
        }
    }
}
