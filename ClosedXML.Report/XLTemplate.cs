﻿using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Report
{
    public class XLTemplate : IDisposable
    {
        private readonly RangeInterpreter _interpreter;
        private bool _disposeWorkbookWithTemplate;

        public bool IsDisposed { get; private set; }

        public IXLWorkbook Workbook { get; private set; }

        static XLTemplate()
        {
            TagsRegister.Add<ColsFitTag>("ColsFit");
            TagsRegister.Add<RowsFitTag>("RowsFit");
            TagsRegister.Add<HiddenTag>("Hidden");
            TagsRegister.Add<HiddenTag>("Hide");
            TagsRegister.Add<OnlyValuesTag>("OnlyValues");
            TagsRegister.Add<AutoFilterTag>("AutoFilter");
            TagsRegister.Add<PageOptionsTag>("PageOptions");
            TagsRegister.Add<ProtectedTag>("Protected");
            TagsRegister.Add<RangeOptionTag>("Range");
            TagsRegister.Add<SortTag>("Sort");
            TagsRegister.Add<SortTag>("Asc");
            TagsRegister.Add<DescTag>("Desc");
            TagsRegister.Add<GroupTag>("Group");
            TagsRegister.Add<PivotTag>("Pivot");
            TagsRegister.Add<FieldPivotTag>("Row");
            TagsRegister.Add<FieldPivotTag>("Column");
            TagsRegister.Add<FieldPivotTag>("Col");
            TagsRegister.Add<FieldPivotTag>("Page");
            TagsRegister.Add<DataPivotTag>("Data");
            TagsRegister.Add<SummaryFuncTag>("SUM");
            TagsRegister.Add<SummaryFuncTag>("AVG");
            TagsRegister.Add<SummaryFuncTag>("AVERAGE");
            TagsRegister.Add<SummaryFuncTag>("COUNT");
            TagsRegister.Add<SummaryFuncTag>("COUNTNUMS");
            TagsRegister.Add<SummaryFuncTag>("MAX");
            TagsRegister.Add<SummaryFuncTag>("MIN");
            TagsRegister.Add<SummaryFuncTag>("PRODUCT");
            TagsRegister.Add<SummaryFuncTag>("STDEV");
            TagsRegister.Add<SummaryFuncTag>("STDEVP");
            TagsRegister.Add<SummaryFuncTag>("VAR");
            TagsRegister.Add<SummaryFuncTag>("VARP");
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
            _interpreter = new RangeInterpreter(null);
        }

        public void Generate()
        {
            CheckIsDisposed();
            foreach (var ws in Workbook.Worksheets.Where(sh => sh.Visibility == XLWorksheetVisibility.Visible && !sh.PivotTables.Any()).ToArray())
            {
                ws.ReplaceCFFormulaeToR1C1();
                _interpreter.Evaluate(ws.AsRange());
                ws.ReplaceCFFormulaeToA1();
            }
        }

        public void AddVariable(object value)
        {
            CheckIsDisposed();
            var type = value.GetType();
            var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance).Where(f => f.IsPublic)
                .Select(f => new { f.Name, val = f.GetValue(value), type = f.FieldType })
                .Concat(type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(f => f.CanRead)
                    .Select(f => new { f.Name, val = f.GetValue(value, new object[] { }), type = f.PropertyType }));

            foreach (var field in fields)
            {
                AddVariable(field.Name, field.val);
            }
        }

        public void AddVariable(string alias, object value)
        {
            CheckIsDisposed();
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
}
