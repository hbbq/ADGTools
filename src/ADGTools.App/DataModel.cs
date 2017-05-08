using ADGTools.Library.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADGTools.App
{
    public class DataModel
    {

        public string date { get; set; }
        public int feeYear { get; set; }
        public List<Person> members { get; set; }

    }
}
