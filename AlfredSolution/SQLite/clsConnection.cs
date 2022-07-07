using System;
using System.Data;
using System.Data.SQLite;
using System.Data.SqlClient;



namespace AlfredCmd
{
    public   class ClsConnection
    {

        public SQLiteConnection cnn = null;
        public SqlConnection ConSql = null;
        public SqlDataReader dr = null;
        public SqlCommand command = null;


        #region "SQLIITE"
        public int fcnSQLiteConnectionOpen(string pathdb)
        {
            try
            {

                cnn = new SQLiteConnection(pathdb);
                cnn.Open();
                return 1;
            }
            catch (SQLiteException ex)
            {
                string code = ex.ErrorCode.ToString();
                return 0;
            }
        }

        public int fcnSQLiteConnectionClose()
        {
            if (cnn.State == ConnectionState.Open)
            {
                cnn.Close();
            }
            return 1;
        }

        public int fcnSQLiteExecuteQuery(string ssql, string param1)
        {
            // executa a query com parametetros.
            SQLiteCommand cmd = new SQLiteCommand();
            cmd.Prepare();
            cmd.Parameters.AddWithValue("@param1", param1);
            try
            {
                int n = cmd.ExecuteNonQuery();
                return n;

            }
            catch (SQLiteException ex)
            {
                throw ex;
            }
            finally
            {
                fcnSQLiteConnectionClose();
            }
        }

        public int fcnSQLiteExecuteQuery2(string ssql)
        {

            // executa a query sem parametros.
            SQLiteCommand cmd = new SQLiteCommand(ssql, cnn);
            cmd.ExecuteNonQuery();
            return 1;
        }


          public int fcnCreateTableSQLite(int i) { 
            
            try
            {
                string strConn = @"Data Source = C:\sqlitedb\sqliteTeste.db";
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
            catch(Exception ex)
            {

                Console.WriteLine(ex.Message, "Erro ao Criar Tabela");               
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
       




        #endregion

        #region "SQLSERVER"

        public SqlConnection fcnSQLConnectionOpen()
        {
            
            ConSql = new SqlConnection("Password=pwb;User ID=userid;Initial Catalog=db;Server=9999999,59070;");
            ConSql.Open();
            return ConSql;
        }

        public void fcnSQLConnectionClose()
        {
            if (ConSql.State == ConnectionState.Open)
            {
                ConSql.Close();
            }
        }

        public SqlDataReader fcnBuscarCap(string strSQL)
        {

            SqlCommand comm = new SqlCommand(strSQL, ConSql);
            SqlDataReader drs = comm.ExecuteReader();
            if (drs.HasRows)
            {
                return drs;

            }
            else
            {
                return null;
            }
        }
        #endregion


    }
}
