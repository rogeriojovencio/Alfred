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
            try
            {
                string sfilePath = spathFile;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook wb = xlApp.Workbooks.Open(@sfilePath);

                xlApp.WindowState = Excel.XlWindowState.xlMaximized;

                if (sVisible == 1)
                {
                    xlApp.Visible = true;
                }
                else
                {
                    xlApp.Visible = false;
                }
                Console.WriteLine("Executando Abertura do Exel...");
                return wb;
            }
            catch (Exception)
            {
                return 0;
                Console.WriteLine("Não foi possível fechar o Arquivo do Exel...");
            }
            
        }
        #endregion OpenWorkbook

        #region CloseWorkbook
        public static object fcnCloseAppExcel(Excel.Workbook wb, int sSaved)
        {
            try
            {
                if (sSaved == 1)
                {
                    wb.Close(1);
                    return 1;
                }
                else
                {
                    wb.Close(0);
                    return 1;
                }
            }
            catch (Exception ex)
            {
                return 0;
                Console.WriteLine($"Não foi possível fechar o Arquivo do Exel...{ex}");
            }



        }

        #endregion CloseWorkbook

        #region fcnLastLine
        public static int fcnLastLine(Excel.Worksheet ws)
        {
            Excel.Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            return rowCount;
        }
        #endregion fcnLastLine

        #region fcnLastColumn
        public static int fcnLastColumn(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range xlRange = ws.UsedRange;
                int colCount = xlRange.Columns.Count;
                return colCount;

            }
            catch (Exception ex)
            {

                Console.WriteLine($"Erro: {ex}");
                return 0;
            }
            
        }
        #endregion fcnLastColumn

        #region fcnGoLastLine
        public static int fcnGoLastLine(Excel.Worksheet ws, int intCol)
        {
            try
            {
                Excel.Range xlRange = ws.UsedRange;
                int rowCount = xlRange.Rows.Count;
                return ws.Cells[rowCount, intCol];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex}");
                return 0;
            }
            
        }
        #endregion fcnGoLastLine

        #region fcnGoLastColumn
        public static int fcnGoLastColumn(Excel.Worksheet ws, int intRow)
        {
            try
            {
                Excel.Range xlRange = ws.UsedRange;
                int colCount = xlRange.Columns.Count;
                return ws.Cells[colCount, intRow];
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Erro: {ex}");
                return 0;
            }
            
        }
        #endregion fcnGoLastColumn

        #region fcnControlCellColor
        public static int fcnControlCellColor(Excel.Worksheet ws, int intRow, int intLastColumn, int intflag, int intColorIndex)
        {
            // metodo tem por objetivo colorir a dimenção da linha e coluna celecionada com a cor desejada evidenciando a linha em questão.

            try
            {
                if (intflag == 1)
                {
                    ws.Cells[intRow, intLastColumn].Interior.ColorIndex = intColorIndex;
                }
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex}");
                return 0;
            }
            
           
        }
        #endregion fcnControlCellColor


        public static int SeekLineClient(Excel.Worksheet ws, string seekString, string sRange )
        {

            seekString = seekString.Trim();
            if (!string.IsNullOrEmpty(seekString))
            {
               

            }

            return 1;

            // continua...
        }



        public static string formatData(string sdata, int stype) 
        {
            string sday;
            string smonth;
            string sYear;
            string sDateOut;
            string sHour;
            string spHour;
            string sminute;
            string sSecond;
            string sTimeOut;


            return "data";
            // continua...
        }

    }

    

}