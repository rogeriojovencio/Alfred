using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using sl = GordonSelenium;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GordonSelenium;
using System.Threading;

namespace AlfredCmd
{
    class Program
    {
        static void Main()
        {

            //-----Testar o codigo dos Projetos

            // abre o browser e 
            //sl.MainBrowsers.OpenBrowser("https://www.google.com");
            FirstScriptTest p = new FirstScriptTest();
            p.ChromeSession();

        }
    }
    [TestClass]
    public class FirstScriptTest
    {
        private bool ret = false;

        [TestMethod]
        public void ChromeSession()
        {
            var driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://google.com");

            var title = driver.Title;
            Assert.AreEqual("Google", title);

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);

            var searchBox = driver.FindElement(By.Name("q"));
            var searchButton = driver.FindElement(By.Name("btnK"));

            searchBox.SendKeys("Selenium");
            searchButton.Click();

            searchBox = driver.FindElement(By.Name("q"));
            var value = searchBox.GetAttribute("value");

            //=============Teste para a Nova classe===================================
            MainBrowsers obj = new MainBrowsers();

            //Abre a pagina. 
            //Localiza se a pagina esta no ar
            //Localiza o campo se existe

            Thread.Sleep(1000);
            //Testando se o objeto existe na tela.
            bool re = obj.SearchTextField(driver, "q", MainBrowsers.SType.Name  );

            // criar um enumerable para definir o tipo de pesquisa utilizando apenas uma função.


            //==========================================================
            Assert.AreEqual("Selenium2", value);

            driver.Quit();
        }

    }


}

