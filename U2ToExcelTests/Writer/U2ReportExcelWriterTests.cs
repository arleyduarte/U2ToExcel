using NUnit.Framework;
using U2ToExcel.Writer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            u2Report =
                new ReportReader().Load(
                    @"C:\Users\zyghtadmin\source\repos\U2ToExcel\U2ToExcel\Resources\REP-ORIGINAL.csv");
        }

        [Test()]
        public void WriteReportTest()
        {

            var destinationPath = $@"TestReport.xlsx";

            var writer = new U2ReportExcelWriter();

            writer.WriteReport(u2Report, destinationPath);

            Assert.IsTrue(File.Exists(destinationPath));
        }
    }
}