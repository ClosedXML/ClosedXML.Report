using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Tests.TestModels;
using ClosedXML.Report.Tests.Utils;
using FluentAssertions;
//using JetBrains.Profiler.Windows.Api;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlsxTemplateTestsBase
    {
        private readonly ITestOutputHelper _output;
        public XlsxTemplateTestsBase(ITestOutputHelper output)
        {
            _output = output;
            LinqToDB.Common.Configuration.Linq.AllowMultipleQuery = true;
        }

        // Because different fonts are installed on Unix,
        // the columns widths after AdjustToContents() will
        // cause the tests to fail.
        // Therefore we ignore the width attribute when running on Unix
        public static bool IsRunningOnUnix
        {
            get
            {
                int p = (int)Environment.OSVersion.Platform;
                return ((p == 4) || (p == 6) || (p == 128));
            }
        }

        protected void XlTemplateTest(string tmplFileName, Action<XLTemplate> arrangeCallback, Action<XLWorkbook> assertCallback)
        {
            /*if (MemoryProfiler.IsActive && MemoryProfiler.CanControlAllocations)
                MemoryProfiler.EnableAllocations();*/

            //MemoryProfiler.Dump();

            var fileName = Path.Combine(TestConstants.TemplatesFolder, tmplFileName);
            var workbook = new XLWorkbook(fileName);
            var template = new XLTemplate(workbook);

            // ARRANGE
            arrangeCallback(template);

            using (var file = new MemoryStream())
            {
                //MemoryProfiler.Dump();
                // ACT
                var start = DateTime.Now;
                template.Generate();
                _output.WriteLine(DateTime.Now.Subtract(start).ToString());
                //MemoryProfiler.Dump();
                workbook.SaveAs(file);
                //MemoryProfiler.Dump();
                file.Position = 0;

                using (var wb = new XLWorkbook(file))
                {
                    // ASSERT
                    assertCallback(wb);
                }
            }
            workbook.Dispose();
            workbook = null;
            template = null;
            GC.Collect();
            //MemoryProfiler.Dump();
        }

        protected void CompareWithGauge(XLWorkbook actual, string fileExpected)
        {
            fileExpected = Path.Combine(TestConstants.GaugesFolder, fileExpected);
            using (var expected = new XLWorkbook(fileExpected))
            {
                actual.Worksheets.Count.ShouldBeEquivalentTo(expected.Worksheets.Count, $"Count of worksheets must be {expected.Worksheets.Count}");

                for (int i = 0; i < actual.Worksheets.Count; i++)
                {
                    WorksheetsAreEqual(expected.Worksheets.ElementAt(i), actual.Worksheets.ElementAt(i), out var messages)
                        .Should().BeTrue(string.Join("," + Environment.NewLine, messages));
                }
            }
        }

        protected bool WorksheetsAreEqual(IXLWorksheet expected, IXLWorksheet actual, out IList<string> messages)
        {
            messages = new List<string>();

            if (expected.Name != actual.Name)
                messages.Add("Worksheet names differ");

            if (expected.RangeUsed().RangeAddress.ToString() != actual.RangeUsed().RangeAddress.ToString())
                messages.Add("Used ranges differ");

            if (expected.Style.ToString() != actual.Style.ToString())
                messages.Add("Worksheet styles differ");

            foreach (var expectedCell in expected.CellsUsed())
            {
                var actualCell = actual.Cell(expectedCell.Address);
                bool cellsAreEqual = true;

                if (actualCell.Value?.ToString() != expectedCell.Value?.ToString())
                {
                    messages.Add($"Cell values are not equal starting from {actualCell.Address}");
                    cellsAreEqual = false;
                }

                if (!string.Equals(actualCell.FormulaA1, expectedCell.FormulaA1, StringComparison.InvariantCultureIgnoreCase))
                {
                    messages.Add($"Cell formulae are not equal starting from {actualCell.Address}");
                    cellsAreEqual = false;
                }

                if (actualCell.DataType != expectedCell.DataType)
                {
                    messages.Add($"Cell data types are not equal starting from {actualCell.Address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.Style.ToString() != actualCell.Style.ToString())
                {
                    messages.Add($"Cell style are not equal starting from {actualCell.Address}");
                    cellsAreEqual = false;
                }

                if (!cellsAreEqual)
                    break; // we don't need thousands of messages
            }

            if (expected.MergedRanges.Count() != actual.MergedRanges.Count())
                messages.Add("Merged ranges counts differ");
            else
            {
                var expectedRanges = expected.MergedRanges
                    .OrderBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                    .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                    .ToList();
                var actualRanges = actual.MergedRanges
                    .OrderBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                    .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                    .ToList();
                for (int i = 0; i < expectedRanges.Count(); i++)
                {
                    var expectedMr = expectedRanges.ElementAt(i);
                    var actualMr = actualRanges.ElementAt(i);
                    if (expectedMr.RangeAddress.ToString() != actualMr.RangeAddress.ToString())
                    {
                        messages.Add($"Merged ranges differ starting from {expectedMr.RangeAddress}");
                        break;
                    }
                }
            }

            if (expected.ConditionalFormats.Count() != actual.ConditionalFormats.Count())
                messages.Add("Conditional format counts differ");

            for (int i = 0; i < expected.ConditionalFormats.Count(); i++)
            {
                var expectedCf = expected.ConditionalFormats.ElementAt(i);
                var actualCf = actual.ConditionalFormats.ElementAt(i);

                if (expectedCf.Range.RangeAddress.ToString() != actualCf.Range.RangeAddress.ToString())
                    messages.Add($"Conditional formats at index {i} have different ranges");

                if (expectedCf.Style.ToString() != actualCf.Style.ToString())
                    messages.Add($"Conditional formats at index {i} have different styles");
            }

            return !messages.Any();
        }
    }
}