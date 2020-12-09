using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using SpreadsheetLight;
using U2ToExcel.Reader;
using Color = System.Drawing.Color;

namespace U2ToExcel.Writer
{
    public class U2ReportExcelWriter
    {
        public  static void WriteReport(U2Report u2Report, string destinationPath)
        {
            var sl = new SLDocument();


            WriteHeader(sl, u2Report);
            SetupColumnsFilter(sl, u2Report);
            WriteBody(sl, u2Report);


            sl.SaveAs(destinationPath);


        }


        private static void SetupColumnsFilter(SLDocument sl, U2Report u2Report)
        {
            var rowStart = u2Report.Headers.Count + 2;
            var cellsContent = ReportReader.GetBodyCells(u2Report.Body[0]);
            var columnCounter = 1;


            foreach (var line in cellsContent)
            {
                sl.SetCellValue(rowStart, columnCounter, line);
                sl.SetCellStyle(rowStart, columnCounter, U2ReportStyles.GetColumnsFilterStyle(sl));
                sl.SetColumnWidth(columnCounter, U2ReportStyles.ColumnWidth);

                columnCounter++;
            }

            var cellRef = $"A{rowStart}";
            var cellRefEnd = $"X{rowStart}";

            sl.Filter(cellRef, cellRefEnd);
        }


        private static void WriteHeader(SLDocument sl, U2Report u2Report)
        {
            var style = sl.CreateStyle();

            style.Font.Bold = true;

            var cellCounter = 1;
            foreach (var line in u2Report.Headers)
            {
                var cellRef = $"A{cellCounter}";
                sl.SetCellValue(cellRef, line);
                sl.SetCellStyle(cellRef, style);
                cellCounter++;
            }
        }



        private static void WriteBody(SLDocument sl, U2Report u2Report)
        {

            var rowStart = u2Report.Headers.Count + 2;

            for (var rowIndex = 1; rowIndex < u2Report.Body.Count; rowIndex++)
            {

                var rowContent = ReportReader.GetBodyCells(u2Report.Body[rowIndex]);

           

                for (var columnIndex = 1; columnIndex < rowContent.Count; columnIndex++)
                {
                    var content = rowContent[columnIndex-1];
                    sl.SetCellValue(rowStart+rowIndex, columnIndex, content);

                    if (rowIndex %  2 == 0)
                        sl.SetCellStyle(rowStart + rowIndex, columnIndex, U2ReportStyles.GetStripStyle(sl));


                }


            }



        }
    }
}