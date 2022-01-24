using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    class cnnExcel
    {

        #region Abrir um Arquivo Excel
        public void fcnOpenAppExcel()
        {
            global.Global g = new global.Global();

            string sPath = @"C:\dados\Files";
            string sFileExcel = @"\Cadastro_prefeituras.xlsm";
            string sfilePath =  sPath + sFileExcel;
            Excel.Application xlApp = new Excel.Application();
           
            Excel.Workbook wb = xlApp.Workbooks.Open(sfilePath);
            Excel.Worksheet ws = wb.Worksheets["Prefeituras"];
            Excel.Worksheet wscnpj = wb.Worksheets["cnpj"];
            
            
            xlApp.WindowState = Excel.XlWindowState.xlMaximized;
            xlApp.DisplayAlerts = false;
            xlApp.Visible = true;
            Console.WriteLine("Executando Abertura do Exel...");

            long LastLine = global.Global.LastRowTotal(ws);



            //    string sEmpresa;
            //    string sUf;
            //    string sFilial;
            //    string sRegional;
            //    string sCidade;
            //    string sCNPJ;
            //    string sRazaoSocial;
            //    string sInscricaoMunicipal;
            //    string sInscricaoEstadual;
            //    string sCodServico;
            //    string sDescricaoServico;
            //    string sAliquotaIss;
            //    string sIdEmpresa;



            //    for (var Line = 0; FirstLine <= LastLine; )
            //    {

            //         sEmpresa = ws.Cells(1, 1);
            //         //sUf;
            //         //sFilial;
            //         //sRegional;
            //         //sCidade;
            //         //sCNPJ;
            //         //sRazaoSocial;
            //         //sInscricaoMunicipal;
            //         //sInscricaoEstadual;
            //         //sCodServico;
            //         //sDescricaoServico;
            //         //sAliquotaIss;
            //         //sIdEmpresa;







            //    }



        }

        #endregion Abrir um Arquivo Excel

        #region Abrir e Ler  e retornar um Arquivo Excel adicionar em uma string.
        public List<string> excelParsing(string fullpath)
        {
            // string data = "";

            List<string> data = new List<string>();
            string der;
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;
               
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fullpath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //either collect data cell by cell or DO you job like insert to DB 
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        data.Add(xlRange.Cells[i, j].Value2.ToString());
                    Console.WriteLine(xlRange.Cells[i, j].Value2.ToString()); ; ;
                    //                    codigo nome    sobrenome

                }
            }

            return data;
        }
        #endregion Abrir e Ler  e retornar um Arquivo Excel adicionar em uma string.



    }
}






   
 