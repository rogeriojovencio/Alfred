
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DAO;
using System.Data;
using FData = AlfredCmd.CnnExcel;
using Gordontxt;

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
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                mes = ws.Name;
                // Case JANEIRO,

                switch (mes)
                {
                    case "JANEIRO":
                        Console.WriteLine($"Trabalhando Planilha mês de: {ws.Name}");
                        fcnDescribeDay(ws, wb);
                        break;
                }
            }
            fcnCloseExcel(wb);
        }
        public static void fcnDescribeDay(Excel.Worksheet ws, Excel.Workbook wb)
        {
            int Line = 1;
            int LastLine = AlfredCmd.CnnExcel.FcnLastLine(ws);
            int LastColumn = 33;
            int Col = 3;
            DAL dt = new DAL();
           // dt.NameFile = "0001FD" + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            //tratamento para os dias do mes
            for (Col = 3; Col < LastColumn; Col++)
            {
                string sDataDia = FData.FormatData(Convert.ToString(ws.Cells[Line, Col].Value), 1);
                fcnEntradasOperacionais(ws, wb, Col, sDataDia);
            }

        }
        public static void fcnEntradasOperacionais(Excel.Worksheet ws, Excel.Workbook wb, int Col, string sDataDia)
        {
            #region Variáveis de Entrada            
            DAL dt = new DAL();            
            int Line = 1;
            int LastLine = AlfredCmd.CnnExcel.FcnLastLine(ws);
            int ColConta = 2;
            int colValues = Col;
            List<object> data = new List<object>();
            WorkFiles wk = new WorkFiles();
            #endregion Variáveis de Entrada

            #region Entradas Operacionais
            // c) Entradas Operacionais
            dt.NameFile = "0001EO" + FData.FormatData(sDataDia, 2) + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            dt.date = FData.FormatData(sDataDia, 1);
            dt.Date_atu = FData.FormatData(DateTime.Now.ToString(), 1).ToString();

            for (Line = 9; Line <= 24; Line++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[Line, ColConta].Value))
                {
                    dt.Name = ws.Cells[Line, ColConta].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = "0";
                    continue;
                }
                dt.date = sDataDia;
                if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells[Line, colValues].Value)))
                {
                    dt.Valor = Convert.ToString(ws.Cells[Line, colValues].Value).Replace(',', '.');
                    // Console.WriteLine($" Este é o valor transformado em decimal: {dt.Valor}");
                }
                else
                {
                    dt.Valor = "0";
                }
                //   Console.WriteLine($"{dt.NameFile};{dt.date};{dt.Name};{dt.Valor}");
                /*
             - Formata data formato sql
             - captura no banco as entradas cadastradas
              */
                data.Add($"{dt.NameFile.ToString()}; {dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()};{dt.Date_atu.ToString()}");
            }

            wk.CreateOrWriteFile("C:\\dados\\Samira_Luciano\\results\\" + dt.NameFile, data, 1000);            

            #endregion Entradas Operacionais
            
            #region Entradas Financeiras
            dt.NameFile = "0001EF" + FData.FormatData(sDataDia, 2) + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            dt.date = FData.FormatData(sDataDia, 1);
            dt.Date_atu = FData.FormatData(DateTime.Now.ToString(),1).ToString();
            for (Line = 27; Line <= 31; Line++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[Line, ColConta].Value))
                {
                    dt.Name = ws.Cells[Line, ColConta].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = "0";
                    continue;
                }
               
                if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells[Line, colValues].Value)))
                {
                    dt.Valor = Convert.ToString(ws.Cells[Line, colValues].Value).Replace(',', '.');
                    Console.WriteLine($" Este é o valor transformado em decimal: {dt.Valor}");
                }
                else
                {
                    dt.Valor = "0";
                }
                Console.WriteLine($"Celula Ativa {dt.NameFile} - {dt.Name}  - {dt.date} - {dt.Valor} - ColConta:{ColConta}");
                /*
             - Formata data formato sql
             - captura no banco as entradas cadastradas
              */
                data.Add($"{dt.NameFile.ToString()}; {dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()};{dt.Date_atu.ToString()}");
            }
            wk.CreateOrWriteFile(("C:\\dados\\Samira_Luciano\\results\\" + dt.NameFile), data, 1000);

            #endregion Entradas Financeiras

            #region Custos com Fornecedores
            dt.NameFile = "0001CF" + FData.FormatData(sDataDia, 2) + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            dt.date = FData.FormatData(sDataDia, 1);
            dt.Date_atu = FData.FormatData(DateTime.Now.ToString(), 1).ToString();
            for (Line = 35; Line <= 46; Line++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[Line, ColConta].Value))
                {
                    dt.Name = ws.Cells[Line, ColConta].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = "0";
                    continue;
                }
                dt.date = DateTime.Now.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells[Line, colValues].Value)))
                {
                    dt.Valor = Convert.ToString(ws.Cells[Line, colValues].Value).Replace(',', '.');
                    Console.WriteLine($" Este é o valor transformado em decimal: {dt.Valor}");
                }
                else
                {
                    dt.Valor = "0";
                }
                //  data.Add($"{dt.NameFile.ToString()}; {Date_atu};{dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()}");
                /*
             - Formata data formato sql
             - captura no banco as entradas cadastradas
              */
                data.Add($"{dt.NameFile.ToString()}; {dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()};{dt.Date_atu.ToString()}");
            }
            wk.CreateOrWriteFile(("C:\\dados\\Samira_Luciano\\results\\" + dt.NameFile), data, 1000);
            #endregion Custos com Fornecedores

            #region Desembolso das Despesas Variávéis
            dt.NameFile = "0001DV" + FData.FormatData(sDataDia, 2) + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            dt.date = FData.FormatData(sDataDia, 1);
            dt.Date_atu = FData.FormatData(DateTime.Now.ToString(), 1).ToString();

            for (Line = 49; Line <= 64; Line++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[Line, ColConta].Value))
                {
                    dt.Name = ws.Cells[Line, ColConta].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = "0";
                    continue;
                }
                dt.date = DateTime.Now.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells[Line, colValues].Value)))
                {
                    dt.Valor = Convert.ToString(ws.Cells[Line, colValues].Value).Replace(',', '.');
                    Console.WriteLine($" Este é o valor transformado em decimal: {dt.Valor}");
                }
                else
                {
                    dt.Valor = "0";
                }
                Console.WriteLine($"Celula Ativa {dt.NameFile} - {dt.Name}  - {dt.date} - {dt.Valor} - ColConta:{ColConta}");
                /*
             - Formata data formato sql
             - captura no banco as entradas cadastradas
              */
                data.Add($"{dt.NameFile.ToString()}; {dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()};{dt.Date_atu.ToString()}");
            }
            wk.CreateOrWriteFile(("C:\\dados\\Samira_Luciano\\results\\" + dt.NameFile), data, 1000);
            #endregion Desembolso das Despesas Variavéis

            #region Desembolso das Despesas Fixas
            dt.NameFile = "0001DF" + FData.FormatData(sDataDia, 2) + DateTime.Now.ToString("yyyyMMddhhmmss").ToString() + ".csv"; // string do tipo do arquivo especifico + 
            dt.date = FData.FormatData(sDataDia, 1);
            dt.Date_atu = FData.FormatData(DateTime.Now.ToString(), 1).ToString();

            for (Line = 67; Line <= 94; Line++)
            {
                if (!string.IsNullOrEmpty(ws.Cells[Line, ColConta].Value))
                {
                    dt.Name = ws.Cells[Line, ColConta].Value;
                }
                else
                {
                    dt.Name = "ÑIdent";
                    dt.Valor = "0";
                    continue;
                }
                dt.date = DateTime.Now.ToString("yyyy-MM-dd");
                if (!string.IsNullOrEmpty(Convert.ToString(ws.Cells[Line, colValues].Value)))
                {
                    dt.Valor = Convert.ToString(ws.Cells[Line, colValues].Value).Replace(',', '.');
                    Console.WriteLine($" Este é o valor transformado em decimal: {dt.Valor}");
                }
                else
                {
                    dt.Valor = "0";
                }
                Console.WriteLine($"Celula Ativa {dt.NameFile} - {dt.Name}  - {dt.date} - {dt.Valor} - ColConta:{ColConta}");
                /*
                - Formata data formato sql
                - captura no banco as entradas cadastradas
                */
                data.Add($"{dt.NameFile.ToString()}; {dt.date.ToString()}; {dt.Name.ToString()}; {dt.Valor.ToString()};{dt.Date_atu.ToString()}");
            }
            wk.CreateOrWriteFile(("C:\\dados\\Samira_Luciano\\results\\" + dt.NameFile), data, 1000);
            #endregion desembolso das Despesas Fixas

        }




        public static void fcnCloseExcel(Excel.Workbook wb)
        {
            string wbname = wb.Name;
            if ((int)AlfredCmd.CnnExcel.FcnCloseWbExcel(wb, 0) == 1)
            {
                Console.WriteLine($"WorkBook encerrado com sucesso! {wbname}");
            }
            AlfredCmd.CnnExcel.FcnCloseAppExcel();
        }

    }
}

