﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace Gordontxt
{
   public  class WorkFiles
    {
        public string  FilePathorigin { get; set; }
        public string Filepathdestiny { get; set; }        

        public void CreateOrWriteFile(string PathFilename, List<object> data, int TotalLines)
        {

            // cria o arquivo            
            StreamWriter sw = new StreamWriter(PathFilename);

            foreach (var item in data)
            {

                sw.WriteLine(item);
                Console.WriteLine($"{item.ToString()}");
            }

            sw.Close();

        }

        
    }
}
