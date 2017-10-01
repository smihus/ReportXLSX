using ClosedXML.Excel;
using System;

namespace ReportXLSX
{
    public class ReportGenerator : IDisposable
    {
        #region Fields
        XLWorkbook _template;
        XLWorkbook _report;
        string _templateFile;
        string _templateWorksheetName;
        string _reportFile;
        #endregion

        public ReportGenerator(string templateFile, string templateWorksheetName)
        {
            _templateFile = templateFile;
            _templateWorksheetName = templateWorksheetName;
        }

        public ReportGenerator(string templateFile, string templateWorksheetName, string reportFile) : this(templateFile, templateWorksheetName)
        {
            _reportFile = reportFile;
        }

        public XLWorkbook Template
        {
            get
            {
                if (_template == null)
                {
                    _template = new XLWorkbook(_templateFile);
                }

                return _template;
            }
            set
            {
                _template = value;
            }
        }

        public XLWorkbook Report
        {
            get
            {
                if (_report == null)
                {
                    _report = new XLWorkbook();
                    _report.AddWorksheet("Report");
                }

                return _report;
            }
            set
            {
                _report = value;
            }
        }

        public void Dispose()
        {
            if (_template != null)
                _template.Dispose();

            if (_report != null)
                _report.Dispose();
        }

        public void InsertRange(string rangeName)
        {
            var templateWS = Template.Worksheet(_templateWorksheetName);
            var table = templateWS.Table("Table1");
            var range = table.HeadersRow();

            var reportWS = Report.Worksheet("Report");
            reportWS.Cell(1, 1).Value = range;
        }

        public void Save()
        {
            Report.SaveAs(_reportFile);
        }

        public void SaveAs(string reportFile)
        {
            Report.SaveAs(reportFile);
        }
    }
}
