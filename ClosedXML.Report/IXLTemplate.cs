using System;
using System.IO;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public interface IXLTemplate : IDisposable
    {
        public IXLWorkbook Workbook { get; }

        public XLGenerateResult Generate();

        public void AddVariable(object value);

        public void AddVariable(string alias, object value);

        public void SaveAs(string file);

        public void SaveAs(string file, SaveOptions options);

        public void SaveAs(string file, bool validate, bool evaluateFormulae = false);

        public void SaveAs(Stream stream);

        public void SaveAs(Stream stream, SaveOptions options);

        public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);
    }
}
