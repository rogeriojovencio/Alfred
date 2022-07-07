using AlfredCmd;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLite
{
   public  class Util
    {

        public int fcnCreateTableSQLite(int i)
        {
            try
            {
                string strConn = @"Data Source = C:\SQLite\sla\gestaosla2.s3db";
                ClsConnection cn = new ClsConnection();
                cn.fcnSQLiteConnectionOpen(strConn);
                string ssqlNew = null;
                switch (i)
                {
                    case 1:
                        ssqlNew = fcnCreateTBL_GESTAO_SLA_ONDA();
                        break;
                    case 2:
                        break;
                    case 4:
                        break;
                }
                cn.fcnSQLiteExecuteQuery2(ssqlNew);
                cn.fcnSQLiteConnectionClose();
                return 1;
            }
            catch
            {

                Console.WriteLine("Erro ao Criar Tabela");               
                return 0;
            }
        }



        public string fcnCreateTBL_GESTAO_SLA_ONDA()
        {
            string ssql = null;



            ssql = "CREATE TABLE IF NOT EXISTS [TBL_GESTAO_SLA_ONDA_] (" +
            "[numero_onda] NUMERIC(18, 0)  UNIQUE NOT NULL," +
            "[descricao_onda] VARCHAR(20)  NULL," +
            "[data_atu] DATE DEFAULT 'GetDate()' NULL" +
             ")";


            return ssql;
        }
    }
}
