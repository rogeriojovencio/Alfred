using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GordonSelenium
{
    public class MainBrowsers
    {
        public static int OpenBrowser(string url)
        {
            try
            {

                var driver = new ChromeDriver();
                driver.Navigate().GoToUrl(url);
                Console.WriteLine($"Url Aberto ok:{url}");
                return 1;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex);
                Console.ReadKey();
                return 0;                
            }
        }

    }
}
