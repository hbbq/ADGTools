using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ADGTools.Library
{
    static class Extensions
    {

        public static string CellValueAsString(this Excel._Worksheet @this, int row, int column)
        {
            return (@this.Cells[row, column] as Excel.Range)?.Value.ToString();
        }

        public static IEnumerable<Models.Restricted.Person> ToRestricted(this IEnumerable<Models.Person> @this)
        {
            var cn = new AutoMapper.MapperConfiguration(c => c.CreateMap<Models.Person, Models.Restricted.Person>());
            var mapper = cn.CreateMapper();
            return @this.Select(o => mapper.Map<Models.Restricted.Person>(o));
        }

    }
}
