using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Menu
{
    class Program
    {
        static void Main(string[] args)
        {
            bool showMenu = true;
           
            while (showMenu)
            {
                showMenu = MainMenu();
            }
           
        }
        private static bool MainMenu()
        {
            Console.Clear();
            Console.WriteLine("\n");
            Console.WriteLine("=============================================================================================+");
            Console.WriteLine("===============================Importar Extratos Bancarios===================================+");
            Console.WriteLine("=============================================================================================+");
            Console.WriteLine("\n");

            Console.WriteLine("\t\tChoose an option:");
            Console.WriteLine("\t\t1) Banco Bradesco");
            Console.WriteLine("\t\t2) Banco Pine");                        
            Console.WriteLine("\t\t22) Exit");
            Console.WriteLine("\n ");
            Console.WriteLine("=============================================================================================+");
            

            Console.Write("\r\n\t\tSelect an option: ");
            #region ConsoleAnt
            switch (Console.ReadLine())
            {
                case "1":
                    Execute("1");
                    return true;
                case "2":
                    Execute("2");
                    return true;
                case "22":                    
                    return false;
                default:
                    return true;
            }
            #endregion ConsoleAnt

        }


        private static string Execute(string t)
        {
            string strExecution = "";

            switch (t)
            {
                case "1":
                    Console.WriteLine("Extrato Banco Bradesco");
                    DisplayResult("Extrato Banco Bradesco");
                    Console.Clear();
                    break;                
                case "2":
                    Console.WriteLine("Extrato Banco Pine");
                    DisplayResult("Extrato Banco Pine ");
                    Console.Clear();
                    break;            
            }

            return strExecution;
        }


        private static void DisplayResult(string message)
        {
            Console.WriteLine($"\r\nJob Executado com sucesso!: {message}");
            Console.Write("\r\nPress Enter to return to Main Menu");
            Console.ReadLine();
        
        }
    }
}
