using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    class Program
    {
        static void Main()
        {

            // Criar Projetos: Class Library
            // Tratamento Excel.....continua...
            // Tratamento com Autoit.
            // projeto Web Spinea MVC.


            /* 0 Objetivo deste procedimento e para prevenir que so poderá efetuar alterações nas planilha 
               que sejam permitidadas pela Aplicação
              passa o Range das planilha que que não serão permitido alterações*/

            //string[] sSheet;
            //sSheet = new string[100];
            //sSheet[0] = "José";
            //sSheet[1] = "Carlos";
            //sSheet[2] = "Macoratti";


            //testando o array
            //foreach(string she in sSheet)
            //{
            //    if (!string.IsNullOrEmpty(she)) { 
            //    Console.WriteLine($"O nome das Planilhas é: {she.ToString()}");
            //    }
            //    else
            //    {
            //        break;
            //    }
            //}







            //sesta o workbook
            Console.WriteLine("Olá Mundo!");
            string strsourceFile = "C:\\dados\\Book1.xlsx";
           // string strsourceFile2 = "C:\\dados\\Entradas.xlsx";

            //Abre o Workbook 
            Excel.Workbook wb1 = (Excel.Workbook)CnnExcel.FcnOpenAppExcel(strsourceFile, 1);
            //Excel.Workbook wb2 = (Excel.Workbook)CnnExcel.FcnOpenAppExcel(strsourceFile2, 1);


            //forma de percorrer uma planilha no Excel
            foreach(Excel.Worksheet ws in wb1.Worksheets)
            {
                if (ws.Name == "José")
                {
                    Console.WriteLine($"Este é o valor da Celula 1 {ws.Cells[1, 1]}");
                }
            }

            //coloca o item dentro da Celula
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);            
            Excel.Worksheet ws1 = (Excel.Worksheet)wb1.Worksheets.get_Item(1);


           //forma de passar parametros para celula evidenciada
            ws1.Cells[1, 1] = "33333";
            Range range = ws1.get_Range("A1");
            Console.WriteLine(range.Value);


            //........................................................................................
            //forma para criar e preencher um array de strings e popula - lo com sas planilhas
            string[] sSheet;
            sSheet = new string[100];
            sSheet[0] = "Menu";
            sSheet[1] = "Auxiliar";
            sSheet[2] = "Config";

            //........................................................................................
            //protegendo e desprotegendo planilhas
            CnnExcel.SuProtecSelectSheets(1, wb1, sSheet); // protege  as planilhas selecionadas.
            CnnExcel.SuProtecSelectSheets(0, wb1, sSheet); // desprotege as planilhas selecionadas.

            //........................................................................................
            //fechando todos os excels abertos, inclusive retirando do task manager.
            CnnExcel.FcnCloseAppExcel();  //Fecha todos os Excels Aberto.





            Console.ReadKey();
        }
    }
}
