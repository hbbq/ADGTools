using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADGTools.Library.Models
{

    public class Person
    {

        [JsonIgnore] public string IdrottsID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        [JsonIgnore] public DateTime BirthDate { get; set; }
        [JsonIgnore] public string Gender { get; set; }
        [JsonIgnore] public string EmailAddress { get; set; }

        [JsonIgnore] public List<Fee> Fees = new List<Fee>();

        public string FullName => $"{FirstName} {LastName}";

        public int Age
        {
            get
            {
                var years = DateTime.Today.Year - BirthDate.Year;
                return years - (DateTime.Today < BirthDate.AddYears(years) ? 1 : 0);
            }
        }

        [JsonIgnore]
        public DateTime? MemberUntil
        {
            get
            {
                var lastPaidFee = Fees.Where(e => e.IsMemberFee && e.IsPaid && e.ForYear.HasValue).OrderBy(e => e.ForYear).LastOrDefault();
                if (lastPaidFee == null) return null;
                return new DateTime(lastPaidFee.ForYear.Value + 1, 4, 30);
            }
        }

        public bool IsPayingMember => MemberUntil.HasValue && MemberUntil.Value >= DateTime.Today;

        public bool IsMember => IsPayingMember || (MemberUntil.HasValue && MemberUntil.Value.AddYears(1) >= DateTime.Today);

    }

}
