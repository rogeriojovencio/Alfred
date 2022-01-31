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

            
            




            Console.WriteLine("Olá Mundo!");
            string strsourceFile = "C:\\dados\\Book1.xlsx";

            Excel.Workbook wb1 = (Excel.Workbook)cnnExcel.fcnOpenAppExcel(strsourceFile,1);
            Excel.Worksheet ws1 = wb1.Sheets[1];
            Console.WriteLine($"Este é o valor da Celula 1 {ws1.Cells[1, 1]}");           

            cnnExcel.fcnCloseAppExcel(wb1, 0);



            Console.ReadKey();
        }
    }
}
