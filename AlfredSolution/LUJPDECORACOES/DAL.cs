using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace LUJPDECORACOES
{
   public  class DAL
    {

        /* propriedades para acoplar os dados vindo do Excel a serem inseridos no anco de daados */

        public string NameFile { get; set; }
        public string Name { get; set; }
        public string date { get; set; }
        public string Valor { get; set; }
        public string Date_atu { get; set; }    
        public List<object> data { get; set; }
        public int TotalLines { get; set; }

        public void CreateOrWriteFile( string PathFilename, List<object> data,  int TotalLines)
        {
            this.TotalLines = TotalLines;
            this.data = data;
            // cria o arquivo
            StreamWriter sw = new StreamWriter(PathFilename);            
            foreach (var item in data)
            {

            }

        }



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
