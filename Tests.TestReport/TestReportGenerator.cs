using ClosedXML.Excel;
using NUnit.Framework;
using ReportXLSX;
using System.Collections.Generic;

namespace Tests.TestReport
{
    [TestFixture]
    public class TestReportGenerator
    {
        ReportGenerator _sut;
        private string _testDirectory;
        List<object> _simpleDataList;

        [SetUp]
        public void SetUp()
        {
            _testDirectory = TestContext.CurrentContext.TestDirectory;




            _simpleDataList = new List<object>();
            _simpleDataList.Add(new { Company = "Company #1", Count = 10, Sum = 1000 });
            _simpleDataList.Add(new { Company = "Company #2", Count = 20, Sum = 2000 });
            _simpleDataList.Add(new { Company = "Company #3", Count = 30, Sum = 3000 });
        }

        [Test]
        public void ShouldHasTemplateProperty()
        {
            var result = _sut.Template;

            Assert.That(result, Is.TypeOf<XLWorkbook>());

            Assert.That(result, Is.Not.Null);
        }

        [Test]
        public void ShouldCreateSimpleReport()
        {
            var templateFile = _testDirectory + @"\Templates\Template.xlsx";
            var templateWorksheetName = "SimpleTemplate";
            var reportFile = @"\SimpleReport.xlsx";

            using (var sut = new ReportGenerator(templateFile, templateWorksheetName))
            {
                var range = "Head";
                sut.InsertRange(range);

                //foreach (var item in _simpleDataList)
                //{
                //    var itemRange = "Row";
                //    sut.InsertRange(itemRange);
                //}

                reportFile = _testDirectory + reportFile;
                sut.SaveAs(reportFile);
            }
        }
    }
}
