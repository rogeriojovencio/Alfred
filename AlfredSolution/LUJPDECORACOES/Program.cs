
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DAL;
using System.Data;

namespace LUJPDECORACOES
{
    class Program
    {

        static void Main(string[] args)
        {
            string dadosConexao = ConfigurationManager.ConnectionStrings["lujpconnection"].ConnectionString;

          // DataAccess dataAccess = new DataAccess();
          // dataAccess.Open();

            DetailWithExcel();

        }


        public static void DetailWithExcel()
            {


            // //sesta o workbook
            // Console.WriteLine("Olá Mundo!");
            string file = "FLUXO CAIXA - LUJP 2022.xlsx";
            string strsource = "C:\\dados\\Samira_Luciano\\";
            string strsourceFile = strsource + file;
            string mes = "";
            string wbname = ""; // variavel serve para armazenar o nome do workbook fechado anteriormente, para efeito de mensagem e ou arquivo de log.

            // //Abre o Workbook             
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)AlfredCmd.CnnExcel.FcnOpenAppExcel(strsourceFile, 1);


             //forma de percorrer uma planilha no Excel
             foreach(Excel.Worksheet ws in wb.Worksheets)
             {       
                mes = ws.Name;
                // Case JANEIRO
                 switch (mes)
                {
                    case "JANEIRO":
                        Console.WriteLine($"Este é o valor da Celula 1 {ws.Name}");
                        break;
                }               
                
             }




            wbname = wb.Name;
            if((int)AlfredCmd.CnnExcel.FcnCloseWbExcel(wb, 0)==1)
            {

                Console.WriteLine($"WorkBook encerrado com sucesso! {wbname}");
            }
            AlfredCmd.CnnExcel.FcnCloseAppExcel();



                        // //coloca o item dentro da Celula
                        // //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);            
                        // Excel.Worksheet ws1 = (Excel.Worksheet)wb1.Worksheets.get_Item(1);


            ////forma de passar parametros para celula evidenciada
            // ws1.Cells[1, 1] = "33333";
            // Range range = ws1.get_Range("A1");
            // Console.WriteLine(range.Value);


            // //........................................................................................
            // //forma para criar e preencher um array de strings e popula - lo com sas planilhas
            // string[] sSheet;
            // sSheet = new string[100];
            // sSheet[0] = "Menu";
            // sSheet[1] = "Auxiliar";
            // sSheet[2] = "Config";

            // //........................................................................................
            // //protegendo e desprotegendo planilhas
            // CnnExcel.SuProtecSelectSheets(1, wb1, sSheet); // protege  as planilhas selecionadas.
            // CnnExcel.SuProtecSelectSheets(0, wb1, sSheet); // desprotege as planilhas selecionadas.

            // //........................................................................................
            // //fechando todos os excels abertos, inclusive retirando do task manager.
            // CnnExcel.FcnCloseAppExcel();  //Fecha todos os Excels Aberto.


        }


    }
}
