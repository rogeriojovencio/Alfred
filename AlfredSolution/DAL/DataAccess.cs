using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAO
{
    public class DataAccess
    {
        #region Objetos Estáticos        
        public static SqlConnection sqlconnection = new SqlConnection();        
        public static SqlCommand comando = new SqlCommand();        
        #endregion

        #region Obter SqlConnection
        public static SqlConnection connection()
        {
            try
            {   
                string dadosConexao = ConfigurationManager.ConnectionStrings["lujpconnection"].ConnectionString;             
                sqlconnection = new SqlConnection(dadosConexao);             
                if (sqlconnection.State == ConnectionState.Closed)
                {                    
                    sqlconnection.Open();
                }                
                return sqlconnection;
            }
            catch (SqlException ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Abre Conexão
        public void Open()
        {
            connection();
        }
        #endregion

        #region Fecha Conexão
        public void Close()
        {
            sqlconnection.Close();
        }
        #endregion

        #region Executar Consulta SQL
        public DataTable ExecutaConsulta(string sql)
        {
            try
            {                
                comando.Connection = connection();                
                comando.CommandText = sql;
                comando.ExecuteScalar();                
                IDataReader dtreader = comando.ExecuteReader();
                DataTable dtresult = new DataTable();
                dtresult.Load(dtreader);                
                sqlconnection.Close();             
                return dtresult;
            }
            catch (Exception ex)
            {            
                throw ex;
            }
        }
        #endregion

        #region Executa uma instrução SQL: INSERT, UPDATE e DELETE
        public int ExecutaAtualizacao(string sql)
        {
            try
            {                
                comando.Connection = connection();
                comando.CommandText = sql;                
                int result = comando.ExecuteNonQuery();
                sqlconnection.Close();                
                return result;
            }
            catch (Exception ex)
            {                
                throw ex;
            }
        }
        #endregion
    }
}
