using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    public static class CnnExcel
    {

        #region Members
        public static object XlApp;
        public static object Wb;
        public static object ws;




        #endregion Members

        #region OpenWorkbook
        public static object FcnOpenAppExcel(string spathFile, int sVisible)
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
                Console.WriteLine("Não foi possível fechar o Arquivo do Exel...");
                return 0;
                
            }

        }
        #endregion OpenWorkbook

        #region CloseWorkbook
        public static object FcnCloseAppExcel(Excel.Workbook wb, int sSaved)
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
                Console.WriteLine($"Não foi possível fechar o Arquivo do Exel...{ex}");
                return 0;
               
            }



        }

        #endregion CloseWorkbook

        #region FcnLastLine
        public static int FcnLastLine(Excel.Worksheet ws)
        {
            Excel.Range xlRange = ws.UsedRange;
            int rowCount = xlRange.Rows.Count;
            return rowCount;
        }
        #endregion FcnLastLine

        #region FcnLastColumn
        public static int FcnLastColumn(Excel.Worksheet ws)
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
        #endregion FcnLastColumn

        #region FcnGoLastLine
        public static int FcnGoLastLine(Excel.Worksheet ws, int intCol)
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
        #endregion FcnGoLastLine

        #region FcnGoLastColumn
        public static int FcnGoLastColumn(Excel.Worksheet ws, int intRow)
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
        #endregion FcnGoLastColumn

        #region fcnControlCellColor
        public static int FcnControlCellColor(Excel.Worksheet ws, int intRow, int intLastColumn, int intflag, int intColorIndex)
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


        public static int SeekLineClient(Excel.Worksheet ws, string seekString, string sRange)
        {

            seekString = seekString.Trim();
            if (!string.IsNullOrEmpty(seekString))
            {


            }

            return 1;

            // continua...
        }



        //public static string FormatData(string sdata, int stype)
        //{
        //    //string sday;
        //    //string smonth;
        //    //string sYear;
        //    //string sDateOut;
        //    //string sHour;
        //    //string spHour;
        //    //string sminute;
        //    //string sSecond;
        //    //string sTimeOut;


        //    return "data";
        //    // continua...
        //}


        public static string[] SuProtecSelectSheets(int stype, Excel.Application xlApp1, Excel.Workbook wb1, string[] she1)
        {
            xlApp1.ScreenUpdating = false;
            xlApp1.DisplayAlerts = false;


            string[] sSheet;
            sSheet = she1;
            int countsht = 0;
            string[] myArray;

            //Filtra no maximo 20 planilhas/ condição imposta pelo sistema.
            myArray = new string[20];

            foreach (string she in sSheet)
            {
                if (!string.IsNullOrEmpty(she))
                {
                    Console.WriteLine($"O nome das Planilhas é: {she}");
                    //pesquisa se existem no workbook atual
                    foreach (Excel.Worksheets ws in wb1.Worksheets)
                    {
                        if(ws.ToString() == she)
                        {
                            // existe entao aplica o metodo protect
                            // soma no array de saida, par retornar os atualizados
                            myArray[countsht] = she.ToString();
                            countsht++;
                            if (countsht > 20)
                            {
                                //retorna as planilhas que conseguiu atualizar
                                return myArray;
                            }
                        }
                    }                    
                }
                else
                {
                    break;
                }
            }

            return myArray;

        }




      
          


        public  static void Protect_Unprotec_sheet(Excel.Worksheet ws,  bool stype)
        {
            if (!stype)
            {
                if (ws.ProtectContents)
                {
                    ws.Unprotect(Password: "!@#");
                }
            }
            else
            {
                if (!ws.ProtectContents)
                {
                    ws.Protect(Password: "!@#", DrawingObjects:true,Contents:true, Scenarios:true , AllowSorting:true, AllowFiltering:true, AllowUsingPivotTables:true);
                }

            }
        }


        



    }
}
 