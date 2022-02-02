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








            Console.WriteLine("Olá Mundo!");
            string strsourceFile = "C:\\dados\\Book1.xlsx";
            Excel.Workbook wb1 = (Excel.Workbook)CnnExcel.FcnOpenAppExcel(strsourceFile, 1);
            Excel.Worksheet ws1 = wb1.Sheets[1];
            Console.WriteLine($"Este é o valor da Celula 1 {ws1.Cells[1, 1]}");


            string[] sSheet;
            sSheet = new string[100];
            sSheet[0] = "José";
            sSheet[1] = "Carlos";
            sSheet[2] = "Macoratti";

          //  cnnExcel.suProtectArrayShhets(sSheet);


            CnnExcel.FcnCloseAppExcel(wb1, 0);



            Console.ReadKey();
        }
    }
}
