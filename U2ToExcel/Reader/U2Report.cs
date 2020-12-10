using System;
using System.Collections.Generic;

namespace U2ToExcel.Reader
{
    public class U2Report
    {
        public List<string> Headers { get; set; }
        public List<string> Body { get; set; }

        public List<string> MoneyColumns { get; set; } = new List<string>();

        public bool IsMoneyColumn(int index)
        {

            var headersBodyCells = ReportReader.GetBodyCells(Body[0]);



            if (index > headersBodyCells.Count)
                return false;

            foreach (var moneyColumn in MoneyColumns)
            {
                if (headersBodyCells[index-1].Equals(moneyColumn))
                {
                    return true;
                }
            }

            return false;
        }
    }
}