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
            WorksheetName = "Report";
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

        public IXLWorksheet ReportWS
        {
            get
            {
                IXLWorksheet ws = null;
                Report.TryGetWorksheet(WorksheetName, out ws);

                if (ws == null)
                    ws = Report.AddWorksheet(WorksheetName);

                return ws;
            }
        }
        public IXLWorksheet TemplateWS
        {
            get
            {
                IXLWorksheet ws = null;
                Template.TryGetWorksheet(_templateWorksheetName, out ws);
                return ws;
            }
        }

        public XLWorkbook Report
        {
            get
            {
                if (_report == null)
                {
                    _report = new XLWorkbook();
                }

                return _report;
            }
            set
            {
                _report = value;
            }
        }

        public string WorksheetName { get; set; }

        public void Dispose()
        {
            if (_template != null)
                _template.Dispose();

            if (_report != null)
                _report.Dispose();
        }

        public void InsertRange(string rangeName)
        {
            var range = TemplateWS.Range(rangeName);

            ReportWS.FirstCell().Value = range;
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
