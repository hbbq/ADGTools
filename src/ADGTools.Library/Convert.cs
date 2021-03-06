﻿using ADGTools.Library.Models;
using System;
using System.Collections.Generic;
using System.Linq;
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

                        var fn = ws.CellValueAsString(row, 3);
                        if (string.IsNullOrEmpty(fn)) break;

                        var ln = ws.CellValueAsString(row, 5);
                        var bd = new DateTime(1901, 1, 1);
                        try
                        {
                            bd = DateTime.ParseExact(ws.CellValueAsString(row, 7), "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch { }
                        var ea = ws.CellValueAsString(row, 11);
                        var gn = ws.CellValueAsString(row, 8) == "Man" ? "M" : "F";
                        var ii = ws.CellValueAsString(row, 6);

                        persons.Add(new Person
                        {
                            BirthDate = bd,
                            EmailAddress = ea,
                            FirstName = fn,
                            Gender = gn,
                            IdrottsId = ii,
                            LastName = ln,
                        });

                    }

                }
                catch
                {
                    // ignored
                }

                wb.Close(false);

            }
            catch
            {
                // ignored
            }

            xl.Quit();

            return persons;

        }

        public static void AddFeesFromExcel(IEnumerable<Person> persons, string fileName)
        {
            var xl = new Excel.Application();
            xl.Visible = true;

            try
            {

                var wb = xl.Workbooks.Open(fileName);

                try
                {

                    var ws = wb.Worksheets[1] as Excel._Worksheet;

                    for (var row = 2; row <= 1000; row++)
                    {

                        var st = ws.CellValueAsString(row, 1);
                        if (string.IsNullOrEmpty(st)) break;

                        var am = int.Parse(ws.CellValueAsString(row, 3).Replace("kr", "").Trim());
                        var dp = DateTime.ParseExact(ws.CellValueAsString(row, 7), "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                        var ds = ws.CellValueAsString(row, 4);
                        var nm = ws.CellValueAsString(row, 2);
                        var ii = ws.CellValueAsString(row, 11);

                        var fee = new Fee
                        {
                            Amount = am,
                            DatePaid = dp,
                            Description = ds,
                            Name = nm,
                            Status = st,
                        };
                        
                        // ReSharper disable once PossibleMultipleEnumeration
                        var pers = persons.FirstOrDefault(e => e.IdrottsId == ii);

                        if (pers != null) pers.Fees.Add(fee);

                    }

                }
                catch
                {
                    // ignored
                }

                wb.Close(false);

            }
            catch
            {
                // ignored
            }

            xl.Quit();
            
        }

        public static IEnumerable<Models.Restricted.Person> PersonsToRestrictedPersons(IEnumerable<Person> persons)
        {
            return persons.ToRestricted();
        }

    }

}
