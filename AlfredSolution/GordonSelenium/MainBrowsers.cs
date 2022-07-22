using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        private static ChromeDriver driver;
        private bool ret = false;

        // Criar um enumerable para definir o tipo de pesquisa utilizando apenas uma função.
        #region SType
        public enum SType
        {
           Id = 1,
           Name = 2,
           Class = 3                
        }
        #endregion SType

        //Open o Browser
        #region OpenBrowser
        public static ChromeDriver OpenBrowser(string url)
        {
            try
            {

                var driver = new ChromeDriver();
                driver.Navigate().GoToUrl(url);
                Console.WriteLine($"Url Aberto ok:{url}");
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex);
                Console.ReadKey();

            }
            return driver;
        }
        #endregion  OpenBrowser

        //Fecha o Browser
        #region CloseBrowser
        public static int CloseBrowser(ChromeDriver driver)
        {
            try
            {
                driver.Quit();
                return 1;
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex);
                Console.ReadKey();
                return 0;
            }
        }
        #endregion CloseBrowser

        //Localiza o elemento na Tela
        #region SearchTextField       
        public bool SearchTextField(ChromeDriver drv, string textfield, SType t)
        {
            ret = false;
            object ElementId;
            try
            {
                switch (t)
                {
                    case SType.Id:
                        ElementId = drv.FindElement(By.Id(textfield));
                        break;
                    case SType.Name:
                        ElementId = drv.FindElement(By.Name(textfield));
                        break;
                    case SType.Class:
                        ElementId = drv.FindElement(By.ClassName(textfield));
                        break;
                }
            }
            catch (Exception)
            {
                ret = false;
            }
            return ret;
        }
        #endregion SearchTextField

        //Clicar no elemento na Tela
        #region SearchTextFieldClick        
        public bool SearchTextFieldClick(ChromeDriver drv, string textfield, SType t)
        {
            ret = false;
            object ElementId;
            try
            {
                switch (t)
                {
                    case SType.Id:
                        drv.FindElement(By.Id(textfield)).Click();
                        break;
                    case SType.Name:
                        drv.FindElement(By.Name(textfield)).Click();
                        break;
                    case SType.Class:
                        drv.FindElement(By.ClassName(textfield)).Click();
                        break;
                }
            }
            catch (Exception)
            {
                ret = false;
            }
            return ret;
        }
        #endregion SearchTextFieldClick

        //Escreve no Elemento Selecionado
        #region SearchTextFieldWrite       
        public bool SearchTextFieldWrite(ChromeDriver drv, string textfield, SType t)
        {
            ret = false;
            object ElementId;
            try
            {
                switch (t)
                {
                    case SType.Id:
                        ElementId = drv.FindElement(By.Id(textfield));
                        break;
                    case SType.Name:
                        ElementId = drv.FindElement(By.Name(textfield));
                        break;
                    case SType.Class:
                        ElementId = drv.FindElement(By.ClassName(textfield));
                        break;
                }
            }
            catch (Exception)
            {
                ret = false;
            }
            return ret;
        }
        #endregion SearchTextFieldWrite

    }
}
