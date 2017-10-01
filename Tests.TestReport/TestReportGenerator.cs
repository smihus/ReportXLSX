using NUnit.Framework;
using ReportXLSX;
using System.Collections.Generic;

namespace Tests.TestReport
{
    [TestFixture]
    public class TestReportGenerator
    {
        ReportGenerator _sut;
        string _testDirectory;
        string _templateFile;
        string _reportFile;
        string _templateWorksheetName;
        List<object> _simpleDataList;

        [SetUp]
        public void SetUp()
        {
            _testDirectory = TestContext.CurrentContext.TestDirectory;

            _templateFile = _testDirectory + @"\Templates\Template.xlsx";
            _templateWorksheetName = "SimpleTemplate";
            _reportFile = _testDirectory + @"\SimpleReport.xlsx";

            _sut = new ReportGenerator(_templateFile, _templateWorksheetName);

            _simpleDataList = new List<object>();
            _simpleDataList.Add(new { Company = "Company #1", Count = 10, Sum = 1000 });
            _simpleDataList.Add(new { Company = "Company #2", Count = 20, Sum = 2000 });
            _simpleDataList.Add(new { Company = "Company #3", Count = 30, Sum = 3000 });
        }

        [TearDown]
        public void TeadDown()
        {
            _sut.Dispose();
        }

        [Test]
        public void ShouldCreateSimpleReport()
        {
            var range = "Head";
            var wsReportName = "Simple company report";
            _sut.WorksheetName = wsReportName;
            _sut.InsertRange(range);

            //foreach (var item in _simpleDataList)
            //{
            //    var itemRange = "Row";
            //    _sut.InsertRange(itemRange, item);
            //}

            var reportFile = _testDirectory + _reportFile;
            _sut.SaveAs(_reportFile);
        }
    }
}
