using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADGTools.Library.Models
{

    public class Fee
    {

        public string Name { get; set; }
        public string Description { get; set; }
        public int Amount { get; set; }
        public string Status { get; set; }
        public DateTime? DatePaid { get; set; }
        public bool IsPaid => DatePaid.HasValue;
        public bool IsMemberFee => Name.Contains("Medlemsavgift");
        public int? ForYear
        {
            get {
                for (var year = 2010; year <= 2100; year++) {
                    if (Name.Contains(year.ToString()) || Description.Contains(year.ToString())) return year;
                }
                return null;
            }

        }

    }

}
