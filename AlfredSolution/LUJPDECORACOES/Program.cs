
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DAO;
using System.Data;

namespace LUJPDECORACOES
{
    class Program
    {

        static void Main(string[] args)
        {
            string dadosConexao = ConfigurationManager.ConnectionStrings["lujpconnection"].ConnectionString;

            //DataAccess dataAccess = new DataAccess();
            //dataAccess.Open();
            
            DetailWithExcel();
            Console.ReadKey();

        }


        public static void DetailWithExcel()
            {


            // //sesta o workbook
            // Console.WriteLine("Olá Mundo!");
            string file = "FLUXO CAIXA - LUJP 2022.xlsx";
            string strsource = "C:\\dados\\Samira_Luciano\\";
            string strsourceFile = strsource + file;
            string mes = "";
            

            // //Abre o Workbook             
            Microsoft.Office.Interop.Excel.Workbook wb = (Microsoft.Office.Interop.Excel.Workbook)AlfredCmd.CnnExcel.FcnOpenAppExcel(strsourceFile, 1);


             //forma de percorrer uma planilha no Excel
             foreach(Excel.Worksheet ws in wb.Worksheets)
             {       
                mes = ws.Name;
                // Case JANEIRO,

                 switch (mes)
                {
                    case "JANEIRO":
                        Console.WriteLine($"Este é o valor da Celula 1 {ws.Name}");
                        fcnEntradasOperacionais(ws, wb);
                        fcnDescribeDay(ws, wb);
                        break;
                }
            }
            fcnCloseExcel(wb);
        }


        public static void fcnDescribeDay(Excel.Worksheet ws, Excel.Workbook wb)
        {
            int Line = 1 ;            
            int LastLine = AlfredCmd.CnnExcel.FcnLastLine(ws);
            int LastColumn = 33;
            int Col = 3;
            DAL dt = new DAL();
            dt.NameFile = "0001FD" + DateTime.Now.ToString("yyyymmddhhmmss").ToString(); // string do tipo do arquivo especifico + 
            //tratamento para os dias do mes
            for (Col = 3; Col < LastColumn; Col++)
            {
                var Cel = ws.Cells[Line, Col].Value;
                Console.WriteLine($"Celula Ativa {Cel}");   
                    /*
                 - Formata data formato sql
                 - captura no banco as entradas cadastradas
                  */
            }
            // saida criar um arquivo csv de linhas e colunas e inserir no banco de dados transacional.
            //outra opção depositar linha a linha no banco de dados, depois fazer a transformação no proprio banco de dados via procedure.

            //tratamento para as colunas das contas






        }

        public static void fcnEntradasOperacionais(Excel.Worksheet ws, Excel.Workbook wb)
        {
            DAL dt = new DAL();
            dt.NameFile = "0001FD" + DateTime.Now.ToString("yyyymmddhhmmss").ToString(); // string do tipo do arquivo especifico + 

            int Line = 1;
            int LastLine = AlfredCmd.CnnExcel.FcnLastLine(ws);
            int LastColumn = 33;
            int Col = 2;


            for (Line = 36; Line < LastLine; Line++)
            {
                var Cel = ws.Cells[Line, Col].Value;

                if(!string.IsNullOrEmpty(ws.Cells[Line, Col].Value))
                {
                    dt.Name = ws.Cells[Line, Col].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = 0;
                    continue;
                }

                dt.date = DateTime.Now.ToString("yyyy-MM-dd");

                if (!string.IsNullOrEmpty(ws.Cells[Line, Col+1].Value))
                {
                    dt.Valor = ws.Cells[Line, Col+ 1].Value;
                }
                else
                {
                    dt.Valor = 0;
                }



                
                Console.WriteLine($"Celula Ativa {dt.NameFile} - {dt.Name}  - {dt.date} - {dt.Valor}");
                /*
             - Formata data formato sql
             - captura no banco as entradas cadastradas
              */
            }



            //NameFile
            //Name
            //date
            //Valor
            //Date_atu








            //"EM DINHEIRO"
            //"DEPOSITO / TRANSFERÊNCIA CTA BRADESCO"
            //"MERCADO PAGO"
            //"MAGALU"
            //"PAGSEGURO"
            //"PAGAR.ME"
            //"PAGHIPER"
            //"SHOPEE"
            //"B2W(SISPAG)"
            //"LEROY MERLIN"
            //"MADEIRA MADEIRA"
            //"AMERICANAS"
            //"WIRECARD"
            //"CLOUD WALK"



        }



    public static void fcnDescribeBill(Excel.Worksheet ws, Excel.Workbook wb)
        {
            int Line = 1;
            int lastLine = AlfredCmd.CnnExcel.FcnLastLine(ws);
            int Col = 2;









        }

        public static void fcnCloseExcel(Excel.Workbook wb)
        {
            string wbname = wb.Name;
            if((int)AlfredCmd.CnnExcel.FcnCloseWbExcel(wb, 0)==1)
            {
                Console.WriteLine($"WorkBook encerrado com sucesso! {wbname}");
            }
            AlfredCmd.CnnExcel.FcnCloseAppExcel();
        }





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
