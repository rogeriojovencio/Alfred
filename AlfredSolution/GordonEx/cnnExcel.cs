using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    class cnnExcel
    {

        #region Members
        public object XlApp;
        public object Wb;
        public object ws;
        #endregion Members

        #region OpenWorkbook
        public object fcnOpenAppExcel(string spath, string sfile)
        {
            string sfilePath = spath + sfile;
            Excel.Application xlApp = new Excel.Application();           
            Excel.Workbook wb = xlApp.Workbooks.Open(sfilePath);
            
            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.DisplayAlerts = false;
            xlApp.Visible = true;
            Console.WriteLine("Executando Abertura do Exel...");
            return wb; 
        }
        #endregion OpenWorkbook

        #region CloseWorkbook
        public object fcnCloseAppExcel( Excel.Workbook wb, int sSaved)
        {
            try
            {
                if (sSaved == 1) { 
                wb.Close(1);
                return 1;
                }
                else
                {
                    wb.Close(0);
                    return 1;
                }
            }
            catch (Exception e)
            {
                return 0;
                Console.WriteLine("Não foi possível fechar o Arquivo do Exel...");
            }


            
        }

        #endregion CloseWorkbook



        // continua...





    }
}






   
 