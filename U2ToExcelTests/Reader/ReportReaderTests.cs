using NUnit.Framework;
using U2ToExcel.Reader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace U2ToExcel.Reader.Tests
{
    [TestFixture()]
    public class ReportReaderTests
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
        public void LoadTest_Header_Successfully()
        {
  


            Assert.IsNotEmpty(u2Report.Headers);

            Assert.AreEqual("Rango Cuentas: 5105 - 52959504", u2Report.Headers[2]);
        }



        [Test()]
        public void LoadTest_Body_Successfully()
        {



            Assert.IsNotEmpty(u2Report.Body);
            Assert.IsTrue( u2Report.Body[0].StartsWith("Cta-Puc"));


        }
    }
}