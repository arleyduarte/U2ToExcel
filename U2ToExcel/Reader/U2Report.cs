using System;
using System.Collections.Generic;

namespace U2ToExcel.Reader
{
    public class U2Report
    {
        public List<string> Headers { get; set; }
        public List<string> Body { get; set; }

        public List<string> MoneyColumns { get; set; }
    }
}