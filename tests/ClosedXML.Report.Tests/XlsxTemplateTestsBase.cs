using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlsxTemplateTestsBase
    {
        protected readonly ITestOutputHelper Output;

        public XlsxTemplateTestsBase(ITestOutputHelper output)
        {
            Output = output;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        }

        protected void XlTemplateTest(string tmplFileName, Action<XLTemplate> arrangeCallback, Action<XLWorkbook> assertCallback)
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, tmplFileName);
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var template = new XLTemplate(stream))
            {
                // ARRANGE
                arrangeCallback(template);

                using (var file = new MemoryStream())
                {
                    // ACT
                    var start = DateTime.Now;
                    template.Generate();
                    Output.WriteLine(DateTime.Now.Subtract(start).ToString());
                    template.SaveAs(file);
                    file.Position = 0;

                    using (var wb = new XLWorkbook(file))
                    {
                        // ASSERT
                        assertCallback(wb);
                    }
                }
            }

            GC.Collect();
        }

        protected void CompareWithGauge(XLWorkbook actual, string fileExpected)
        {
            fileExpected = Path.Combine(TestConstants.GaugesFolder, fileExpected);
            if (!File.Exists(fileExpected))
            {
                actual.SaveAs(Path.Combine("Output", Path.GetFileName(fileExpected)));
                throw new FileNotFoundException("Gauge file not found.", fileExpected);
            }

            using (var expectStream = File.Open(fileExpected, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var expected = new XLWorkbook(expectStream))
            {
                var shEquals = actual.Worksheets.Count == expected.Worksheets.Count;
                if (!shEquals)
                    actual.SaveAs(Path.Combine("Output", Path.GetFileName(fileExpected)));
                shEquals.Should().BeTrue($"Count of worksheets must be {expected.Worksheets.Count}");

                for (int i = 0; i < actual.Worksheets.Count; i++)
                {
                    var areEqual = WorksheetsAreEqual(expected.Worksheets.ElementAt(i), actual.Worksheets.ElementAt(i), out var messages);

                    if (!areEqual)
                        actual.SaveAs(Path.Combine("Output", Path.GetFileName(fileExpected)));

                    areEqual.Should().BeTrue(string.Join("," + Environment.NewLine, messages));
                }
            }
        }

        protected bool WorksheetsAreEqual(IXLWorksheet expected, IXLWorksheet actual, out IList<string> messages)
        {
            messages = new List<string>();

            if (expected.Name != actual.Name)
                messages.Add("Worksheet names differ");

            if (expected.RangeUsed()?.RangeAddress?.ToString() != actual.RangeUsed()?.RangeAddress?.ToString())
                messages.Add("Used ranges differ");

            if (expected.Style.ToString() != actual.Style.ToString())
                messages.Add("Worksheet styles differ");

            if (!expected.PageSetup.RowBreaks.All(actual.PageSetup.RowBreaks.Contains)
                || expected.PageSetup.RowBreaks.Count != actual.PageSetup.RowBreaks.Count)
                messages.Add("PageBreaks differ");

            if (expected.PageSetup.PagesTall != actual.PageSetup.PagesTall)
                messages.Add("PagesTall differ");

            if (expected.PageSetup.PagesWide != actual.PageSetup.PagesWide)
                messages.Add("PagesWide differ");

            if (expected.PageSetup.PageOrientation != actual.PageSetup.PageOrientation)
                messages.Add("PageOrientation differ");

            if (expected.PageSetup.PageOrder != actual.PageSetup.PageOrder)
                messages.Add("PageOrder differ");

            var usedCells = expected.CellsUsed(XLCellsUsedOptions.All).Select(c => c.Address)
                .Concat(actual.CellsUsed(XLCellsUsedOptions.All).Select(c => c.Address))
                .Distinct();
            foreach (var address in usedCells)
            {
                var expectedCell = expected.Cell(address);
                var actualCell = actual.Cell(address);
                bool cellsAreEqual = true;

                if (!expectedCell.HasFormula && !actualCell.HasFormula && actualCell.Value.ToString() != expectedCell.Value.ToString())
                {
                    messages.Add($"Cell values are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (!string.Equals(actualCell.FormulaA1, expectedCell.FormulaA1, StringComparison.InvariantCultureIgnoreCase))
                {
                    messages.Add($"Cell formulae are not equal starting from {address}. Actual: '{actualCell.FormulaA1}'; Expected: '{expectedCell.FormulaA1}'");
                    cellsAreEqual = false;
                }

                if (!expectedCell.HasFormula && actualCell.DataType != expectedCell.DataType &&
                    actualCell.DataType != XLDataType.Blank && expectedCell.GetValue<string>() != "") // Special case for treating blanks as equal to empty strings for not changing all gauges created in older versions (with no "blank" type)
                {
                    messages.Add($"Cell data types are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.HasComment != actualCell.HasComment
                    || (expectedCell.HasComment && !Equals(expectedCell.GetComment(), actualCell.GetComment())))
                {
                    messages.Add($"Cell comments are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.HasHyperlink != actualCell.HasHyperlink
                    || (expectedCell.HasHyperlink && !Equals(expectedCell.GetHyperlink(), actualCell.GetHyperlink())))
                {
                    messages.Add($"Cell Hyperlink are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.HasRichText != actualCell.HasRichText
                    || (expectedCell.HasRichText && !expectedCell.GetRichText().Equals(actualCell.GetRichText())))
                {
                    messages.Add($"Cell RichText are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.HasDataValidation != actualCell.HasDataValidation
                    || (expectedCell.HasDataValidation && !Equals(expectedCell.GetDataValidation(), actualCell.GetDataValidation())))
                {
                    messages.Add($"Cell DataValidation are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (expectedCell.Style.ToString() != actualCell.Style.ToString())
                {
                    messages.Add($"Cell style are not equal starting from {address}");
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
                for (int i = 0; i < expectedRanges.Count; i++)
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
            else
            {
                var expectedFormats = expected.ConditionalFormats
                    .OrderBy(r => r.Range.RangeAddress.FirstAddress.ColumnNumber)
                    .ThenBy(r => r.Range.RangeAddress.FirstAddress.RowNumber)
                    .ToList();
                var actualFormats = actual.ConditionalFormats
                    .OrderBy(r => r.Range.RangeAddress.FirstAddress.ColumnNumber)
                    .ThenBy(r => r.Range.RangeAddress.FirstAddress.RowNumber)
                    .ToList();

                for (int i = 0; i < expectedFormats.Count; i++)
                {
                    var expectedCf = expectedFormats.ElementAt(i);
                    var actualCf = actualFormats.ElementAt(i);

                    if (expectedCf.Range.RangeAddress.ToString() != actualCf.Range.RangeAddress.ToString())
                        messages.Add($"Conditional formats actual {actualCf.Range.RangeAddress}, but expected {expectedCf.Range.RangeAddress}.");

                    if (expectedCf.Style.ToString() != actualCf.Style.ToString())
                        messages.Add($"Conditional formats at {actualCf.Range.RangeAddress} have different styles");

                    if (expectedCf.Values.Count != actualCf.Values.Count)
                        messages.Add($"Conditional formats at {actualCf.Range.RangeAddress} counts differ");

                    for (int j = 1; j <= expectedCf.Values.Count; j++)
                    {
                        if (expectedCf.Values[j].Value != actualCf.Values[j].Value)
                        {
                            messages.Add($"Conditional formats at {actualCf.Range.RangeAddress} have different values");
                            break;
                        }
                    }
                }
            }

            return !messages.Any();
        }

        private bool Equals(XLHyperlink expectedHyperlink, XLHyperlink actualHyperlink)
        {
            if (expectedHyperlink == actualHyperlink) return true;

            return expectedHyperlink.IsExternal == actualHyperlink.IsExternal
                   && expectedHyperlink.ExternalAddress == actualHyperlink.ExternalAddress
                   && expectedHyperlink.InternalAddress == actualHyperlink.InternalAddress;
        }

        private bool Equals(IXLComment expectedComment, IXLComment actualComment)
        {
            if (expectedComment == actualComment) return true;

            return // TODO expectedComment.Equals(actualComment) // ClosedXML issue #1450
                    expectedComment.Text == actualComment.Text
                   /*&& expectedComment.Style!!! == actualComment.Style!!!
                   && expectedComment.Position.Column == actualComment.Position.Column
                   && expectedComment.Position.ColumnOffset == actualComment.Position.ColumnOffset
                   && expectedComment.Position.Row == actualComment.Position.Row
                   && expectedComment.Position.RowOffset == actualComment.Position.RowOffset
                   && expectedComment.ZOrder == actualComment.ZOrder*/
                   && expectedComment.ShapeId == actualComment.ShapeId
                   /*&& expectedComment.Visible == actualComment.Visible*/;
        }

        private bool Equals(IXLDataValidation expectedValidation, IXLDataValidation actualValidation)
        {
            if (expectedValidation == actualValidation) return true;

            return expectedValidation.IgnoreBlanks == actualValidation.IgnoreBlanks
                   && expectedValidation.Ranges.ToString() == actualValidation.Ranges.ToString()
                   && expectedValidation.InCellDropdown == actualValidation.InCellDropdown
                   && expectedValidation.ShowErrorMessage == actualValidation.ShowErrorMessage
                   && expectedValidation.ShowInputMessage == actualValidation.ShowInputMessage
                   && expectedValidation.InputTitle == actualValidation.InputTitle
                   && expectedValidation.InputMessage == actualValidation.InputMessage
                   && expectedValidation.ErrorTitle == actualValidation.ErrorTitle
                   && expectedValidation.ErrorMessage == actualValidation.ErrorMessage
                   && expectedValidation.ErrorStyle == actualValidation.ErrorStyle
                   && expectedValidation.AllowedValues == actualValidation.AllowedValues
                   && expectedValidation.Operator == actualValidation.Operator
                   && expectedValidation.MinValue == actualValidation.MinValue
                   && expectedValidation.MaxValue == actualValidation.MaxValue;
        }
    }
}
