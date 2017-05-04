using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADGTools.App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ps = Library.Convert.ExcelToPersons(@"C:\Users\henber\Downloads\ExportedPersons (1).xls");
            Library.Convert.AddFeesFromExcel(ps, @"C:\Users\henber\Downloads\ExportFile (4).xls");
            foreach(var p in ps.Where(o => o.IsPayingMember))
            {
                System.Diagnostics.Debug.WriteLine(p.FullName);
            }
        }
    }
}
