using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AlfredCmd
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Olá Mundo!");

            cnnExcel obj = new cnnExcel();            
            obj.excelParsing(@"C:\dados\Book1.xlsm");
            Console.ReadKey();
        }
    }
}
