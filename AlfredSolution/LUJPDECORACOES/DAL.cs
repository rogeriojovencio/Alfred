using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LUJPDECORACOES
{
   public  class DAL
    {

        /* propriedades para acoplar os dados vindo do Excel a serem inseridos no anco de daados */

        public string NameFile { get; set; }
        public string Name { get; set; }
        public string date { get; set; }
        public Double Valor { get; set; }
        public DateTime Date_atu { get; set; }        


        public int fcnInsertTableImport()
        {            
            try
            {
                Console.WriteLine("Inserido com sucesso!");
                return 1;
            }
            catch (Exception)
            {
                Console.WriteLine("Erro ao Inserir o registro!");
                throw;
            }
        }
   }

}
