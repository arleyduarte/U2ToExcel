using System.Collections.Generic;
using System.IO;
using System.Text;

namespace U2ToExcel.Reader
{
    public  class ReportReader
    {
        public static U2Report  Load(string path, string moneyColumns = "")
        {

            var u2Report = new U2Report();
            var lines = File.ReadAllLines(path, Encoding.Latin1);

            u2Report.Headers = GetHeader(lines);
            u2Report.Body = GetBody(lines);
            u2Report.MoneyColumns = ArgumentHelper.GetMoneyColumns(moneyColumns);

            return u2Report;
        }


        private static List<string> GetHeader(string[] lines)
        {
            var headers = new List<string>();

            foreach (var line in lines)
            {

                if(IsBodyStart(line))
                    break;

                headers.Add(line.Split(";")[0].Trim());

            }


            return headers;
        }

        private static bool IsBodyStart(string line)
        {
            return line.StartsWith("@") || line.StartsWith("Cta");
        }



        private static List<string> GetBody(string[] lines)
        {
            var body = new List<string>();

            var isBody = false;

            foreach (var line in lines)
            {

                if (IsBodyStart(line))
                {
                    isBody = true;

                }

                if (isBody)
                {
                    if(!IsEmptyLine(line))
                     body.Add(line);
                }

            }



            return body;
        }

        private static bool IsEmptyLine(string line)
        {
            var cells = GetBodyCells(line);
            foreach (var word in cells)
            {
                if (!string.IsNullOrWhiteSpace(word))
                {
                    return false;
                }
            }

            return true;
        }


        public static List<string> GetBodyCells(string line)
        {
            var cells = new List<string>();
            var words = line.Split(';');

            foreach (var word in words)
            {

                var cleaned = word.Replace("\0", string.Empty);
                cells.Add(cleaned.Trim());

            }
            return cells;
        }
    }
}
