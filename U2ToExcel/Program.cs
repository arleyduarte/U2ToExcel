using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Presentation;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;
using log4net;
using log4net.Config;
using U2ToExcel.Reader;
using U2ToExcel.Writer;

namespace U2ToExcel
{
    class Program
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        //https://zyght.blob.core.windows.net/acorde-demo/net5.0.zip
        static void Main(string[] args)
        {

            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));








            Log.Info($"Start 3 ------------------------");
            Log.Info(args.Length);
            Log.Info(args);

            var reportArg = GetArguments(args);


            if (string.IsNullOrWhiteSpace(reportArg.Origin) || string.IsNullOrWhiteSpace(reportArg.Destination))
            {
                Log.Error("Debe ingresar Nombre del informe origina (ruta), Nombre del informe final en el Excel ");

                return;
            }

            Log.Info($"Archivo origen: {reportArg.Origin}");
            Log.Info($"Archivo destino: {reportArg.Destination}");

            try
            {


                var u2Report = ReportReader.Load(reportArg.Origin);
                Log.Info($"Lineas cargadas: {u2Report.Body.Count}");

                U2ReportExcelWriter.WriteReport(u2Report, reportArg.Destination);
            }
            catch (Exception e)
            {
                Log.Error(e);
            }





            Log.Info($"END ------------------------");



        }


        private static ReportArguments GetArguments(string[] args)
        {
            var reportArg = new ReportArguments();


            for (int i = 0; i < args.Length; i++)
            {
                switch (i)
                {
                    case 0:
                        reportArg.Origin = args[i];
                        break;
                    case 1:
                        reportArg.Destination = args[i];
                        break;
                }
            }


            /*
            foreach (var argument in args)
            {
                Log.Info($"GetArguments args  {argument}");
                var all = argument.Split(',');

                for (int i = 0; i < all.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            reportArg.Origin = all[i];
                            break;
                        case 1:
                            reportArg.Destination = all[i];
                            break;
                    }
                }

            }
            */




            return reportArg;


        }
        private class ReportArguments
        {
            public string Origin {get; set; }
            public string Destination { get; set; }
        }


    }
}
