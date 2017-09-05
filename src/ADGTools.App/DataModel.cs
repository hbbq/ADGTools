using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace ADGTools.App
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class DataModel<T>
    {

        public string date { get; set; }
        public int feeYear { get; set; }
        public List<T> members { get; set; }

    }
}
