using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Report
{
    public class XLTemplate : IDisposable
    {
        private readonly RangeInterpreter _interpreter;
        private readonly bool _disposeWorkbookWithTemplate;
        private readonly TemplateErrors _errors = new TemplateErrors();

        public bool IsDisposed { get; private set; }

        public IXLWorkbook Workbook { get; private set; }

        static XLTemplate()
        {
            TagsRegister.Add<RangeOptionTag>("Range", 255);
            TagsRegister.Add<RangeOptionTag>("SummaryAbove", 255);
            TagsRegister.Add<RangeOptionTag>("DisableGrandTotal", 255);
            TagsRegister.Add<GroupTag>("Group", 200);
            TagsRegister.Add<PivotTag>("Pivot", 180);
            TagsRegister.Add<FieldPivotTag>("Row", 180);
            TagsRegister.Add<FieldPivotTag>("Column", 180);
            TagsRegister.Add<FieldPivotTag>("Col", 180);
            TagsRegister.Add<FieldPivotTag>("Page", 180);
            TagsRegister.Add<DataPivotTag>("Data", 180);
            TagsRegister.Add<SortTag>("Sort", 128);
            TagsRegister.Add<SortTag>("Asc", 128);
            TagsRegister.Add<DescTag>("Desc", 128);
            TagsRegister.Add<SummaryFuncTag>("SUM", 50);
            TagsRegister.Add<SummaryFuncTag>("AVG", 50);
            TagsRegister.Add<SummaryFuncTag>("AVERAGE", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNT", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNTA", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNTNUMS", 50);
            TagsRegister.Add<SummaryFuncTag>("MAX", 50);
            TagsRegister.Add<SummaryFuncTag>("MIN", 50);
            TagsRegister.Add<SummaryFuncTag>("PRODUCT", 50);
            TagsRegister.Add<SummaryFuncTag>("STDEV", 50);
            TagsRegister.Add<SummaryFuncTag>("STDEVP", 50);
            TagsRegister.Add<SummaryFuncTag>("VAR", 50);
            TagsRegister.Add<SummaryFuncTag>("VARP", 50);
            TagsRegister.Add<OnlyValuesTag>("OnlyValues", 40);
            TagsRegister.Add<AutoFilterTag>("AutoFilter", 10);
            TagsRegister.Add<ColsFitTag>("ColsFit", 0);
            TagsRegister.Add<RowsFitTag>("RowsFit", 0);
            TagsRegister.Add<HiddenTag>("Hidden", 0);
            TagsRegister.Add<HiddenTag>("Hide", 0);
            TagsRegister.Add<PageOptionsTag>("PageOptions", 0);
            TagsRegister.Add<ProtectedTag>("Protected", 0);
            TagsRegister.Add<DeleteTag>("Delete", 0);
        }

        public XLTemplate(string fileName) : this(new XLWorkbook(fileName))
        {
            _disposeWorkbookWithTemplate = true;
        }

        public XLTemplate(Stream stream) : this(new XLWorkbook(stream))
        {
            _disposeWorkbookWithTemplate = true;
        }

        public XLTemplate(IXLWorkbook workbook)
        {
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook), "Workbook cannot be null");
            _interpreter = new RangeInterpreter(null, _errors);
        }

        public XLGenerateResult Generate()
        {
            CheckIsDisposed();
            foreach (var ws in Workbook.Worksheets.Where(sh => sh.Visibility == XLWorksheetVisibility.Visible && !sh.PivotTables.Any()).ToArray())
            {
                ws.ReplaceCFFormulaeToR1C1();
                _interpreter.Evaluate(ws.AsRange());
                ws.ReplaceCFFormulaeToA1();
            }
            return new XLGenerateResult(_errors);
        }

        public void AddVariable(object value)
        {
            CheckIsDisposed();
            if (value is IDictionary dictionary)
            {
                foreach (DictionaryEntry entry in dictionary)
                {
                    AddVariable(entry.Key.ToString(), entry.Value);
                }
            }
            else
            {
                var type = value.GetType();
                var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance).Where(f => f.IsPublic)
                    .Select(f => new {f.Name, val = f.GetValue(value), type = f.FieldType})
                    .Concat(type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(f => f.CanRead)
                        .Select(f => new {f.Name, val = f.GetValue(value, new object[] { }), type = f.PropertyType}));

                foreach (var field in fields)
                {
                    AddVariable(field.Name, field.val);
                }
            }
        }

        public void AddVariable(string alias, object value)
        {
            CheckIsDisposed();
            if (value is DataTable)
                value = ((DataTable) value).Rows.Cast<DataRow>();
            _interpreter.AddVariable(alias, value);
        }

        public void SaveAs(string file)
        {
            CheckIsDisposed();
            Workbook.SaveAs(file);
        }

        public void SaveAs(Stream stream)
        {
            CheckIsDisposed();
            Workbook.SaveAs(stream);
        }

        public void Dispose()
        {
            if (IsDisposed)
                return;

            if (_disposeWorkbookWithTemplate)
                Workbook.Dispose();
            Workbook = null;
            IsDisposed = true;
        }

        private void CheckIsDisposed()
        {
            if (IsDisposed)
                throw new ObjectDisposedException("Template has been disposed");
        }
    }

    public class XLGenerateResult
    {
        public XLGenerateResult(TemplateErrors errors)
        {
            ParsingErrors = errors;
        }

        public bool HasErrors => ParsingErrors.Count > 0;
        public TemplateErrors ParsingErrors { get; }
    }
}
