using NUnit.Framework;
using U2ToExcel.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;
using U2ToExcel.Reader;

namespace U2ToExcel.Writer.Tests
{
    [TestFixture()]
    public class U2ReportExcelWriterTests
    {

        private U2Report u2Report;

        [SetUp()]
        public void SetUp()
        {
            /* u2Report =
                  ReportReader.Load(
                     @"C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\REP-ORIGINAL.csv");*/

         //   var columns = "Debitos:Saldo:Creditos:Base";
         string columns = null;

            u2Report =
                ReportReader.Load(
                    @"C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\REP-ORIGINAL-Short.csv", columns);
        }

        [Test()]
        public void WriteReportTest()
        {

            var destinationPath = $@"TestReportShort.xlsx";

   

            U2ReportExcelWriter.WriteReport(u2Report, destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));




        }

        [Test()]
        public  void WriteReport_Functions()
        {
            var sl = new SLDocument();
            var destinationPath = $@"TestReportF.xlsx";
            var cellRef = $"A1";


            var style = new SLStyle();
            style.FormatCode = "#,##0.00";
            sl.SetCellStyle("A1", style);

            sl.SetCellValue(cellRef,  148100.98);

            sl.SaveAs(destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));
        }
    }



}