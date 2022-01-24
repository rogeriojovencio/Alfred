using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    public static class cnnExcel
    {

        #region Members
        public static object XlApp;
        public static object Wb;
        public static object ws;
        #endregion Members

        #region OpenWorkbook
        public static object fcnOpenAppExcel(string spathFile, int sVisible)
        {
            string sfilePath = spathFile;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook wb = xlApp.Workbooks.Open(@sfilePath);

            xlApp.WindowState = Excel.XlWindowState.xlMaximized;

            if (sVisible == 1) {
                xlApp.Visible = true;
            }
            else
            {
                xlApp.Visible = false;
            }
            Console.WriteLine("Executando Abertura do Exel...");
            return wb;
        }
        #endregion OpenWorkbook

        #region CloseWorkbook
        public static object fcnCloseAppExcel(Excel.Workbook wb, int sSaved)
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

        #region fcnLastLine
        public static int fcnLastLine(Excel.Worksheet ws)
        {            
            Excel.Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            return rowCount;
        }
        #endregion fcnLastLine

        #region fcnLastColumn
        public static int fcnLastColumn(Excel.Worksheet ws)
        {
            Excel.Range xlRange = ws.UsedRange;            
            int colCount = xlRange.Columns.Count;
            return colCount;
        }
        #endregion fcnLastColumn




    }

    // continua...

}









   
 