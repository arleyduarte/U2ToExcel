using System;
using System.Collections.Generic;
using System.Reflection;
using log4net;

namespace U2ToExcel
{
    public static  class ArgumentHelper
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public static List<string> GetMoneyColumns(string args)
        {
            var columns = args.Split(":");
            var rColumns = new List<string>();
            foreach (var column in columns)
            {
                if(!string.IsNullOrWhiteSpace(column))
                    rColumns.Add(column);
            }

            return rColumns;
        }



        public static ReportArguments GetArguments(string[] argsV)
        {
            var reportArg = new ReportArguments();

            var args = SplitArgument(argsV);

            for (var i = 0; i < args.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        reportArg.Origin = args[i];
                        break;
                    case 1:
                        reportArg.Destination = args[i];
                        break;

                    case 2:
                        reportArg.MoneyColumns = args[i];
                        break;
                }
            }

            return reportArg;


        }

        private static List<string> SplitArgument(string[] args)
        {




            var arguments = new List<string>();
            if (args.Length == 1)
            {

                var all = args[0].Split(' ');

                foreach (var arg in all)
                {
                    arguments.Add(arg);
                }

                return arguments;
            }


            foreach (var arg in args)
            {
                arguments.Add(arg);
            }


            return arguments;
        }
    }
}