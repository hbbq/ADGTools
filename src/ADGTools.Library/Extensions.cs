using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ADGTools.Library
{
    static class Extensions
    {

        public static string CellValueAsString(this Excel._Worksheet @this, int row, int column)
        {
            return (@this.Cells[row, column] as Excel.Range).Value.ToString();
        }

    }
}
