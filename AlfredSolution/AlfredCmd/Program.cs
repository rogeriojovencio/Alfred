﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfredCmd
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Olá Mundo!");
            string strsourceFile = "C:\\dados\\Book1.xlsx";

            Excel.Workbook wb1 = (Excel.Workbook)cnnExcel.fcnOpenAppExcel(strsourceFile,1);
            Excel.Worksheet ws1 = wb1.Sheets[1];

            Excel.Range xlRange = ws1.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            cnnExcel.fcnCloseAppExcel(wb1, 0);



            Console.ReadKey();
        }
    }
}
