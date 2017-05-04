using ADGTools.Library.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ADGTools.Library
{

    public static class Convert
    {

        public static List<Person> ExcelToPersons(string fileName)
        {

            var persons = new List<Person>();

            var xl = new Excel.Application();
            xl.Visible = true;

            try
            {

                var wb = xl.Workbooks.Open(fileName);

                try
                {

                    var ws =  wb.Worksheets[1] as Excel._Worksheet;                    

                    for(var row = 2; row <= 1000; row++)
                    {

                        var fn = (ws.Cells[row, 3] as Excel.Range).Value as string;
                        if (string.IsNullOrEmpty(fn)) break;

                        var ln = (ws.Cells[row, 5] as Excel.Range).Value as string;
                        var bd = DateTime.ParseExact((ws.Cells[row, 7] as Excel.Range).Value as string, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                        var ea = (ws.Cells[row, 11] as Excel.Range).Value as string;
                        var gn = ((ws.Cells[row, 8] as Excel.Range).Value as string) == "Man" ? "M" : "F";
                        var ii = (ws.Cells[row, 6] as Excel.Range).Value as string;

                        persons.Add(new Person
                        {
                            BirthDate = bd,
                            EmailAddress = ea,
                            FirstName = fn,
                            Gender = gn,
                            IdrottsID = ii,
                            LastName = ln,
                        });

                    }

                }
                catch { }

                wb.Close(false);

            }
            catch { }

            xl.Quit();

            return persons;

        }

    }

}
