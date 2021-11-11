using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AlfredCmd.global
{
    class Global
    {
        // funções auxiliares.          
       /// <summary>
       /// 
       /// Esta função tem por objetivo de retornar a ultima linha preenchida da  planilha.
       /// </summary>
       /// <param name="ws"></param>
       /// <returns></returns>       
       public static int LastRowTotal(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            
            Microsoft.Office.Interop.Excel.Range lastCell = ws.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            return lastCell.Row;
        }




    }
}
