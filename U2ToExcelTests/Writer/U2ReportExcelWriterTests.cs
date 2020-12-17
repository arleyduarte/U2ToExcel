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

        private string resourcesPath = @"C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\";

        [SetUp()]
        public void SetUp()
        {

            var columns = "Debitos:Saldo:Creditos:Base";


            u2Report =
                ReportReader.Load(
                    @"C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\CP0220RU2.csv", columns);
        }

        [Test()]
        public void WriteReportTest()
        {

            var destinationPath = $@"TestReportShort.xlsx";
            U2ReportExcelWriter.WriteReport(u2Report, destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));

        }

        [Test()]
        public void WriteWalletTest()
        {

            var columns = "Plazo:Dias-V:Corriente:V 1-30:V 31-60:V 61-90:V 91-120:V 121-180:V 181-360:V 360-mas:TOTAL                                                                                                                   ";

            u2Report =
                ReportReader.Load($"{resourcesPath}CARTERA-U2.csv", columns);


            var destinationPath = $@"Catera.xlsx";
            U2ReportExcelWriter.WriteReport(u2Report, destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));

        }



        [Test()]
        public void WriteReport_Functions()
        {
            var sl = new SLDocument();
            var destinationPath = $@"TestReportF.xlsx";
            var cellRef = $"A1";


            var style = new SLStyle();
            style.FormatCode = "#,##0.00";
            sl.SetCellStyle("A1", style);

            sl.SetCellValue(cellRef, 148100.98);

            sl.SaveAs(destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));
        }
    }



}