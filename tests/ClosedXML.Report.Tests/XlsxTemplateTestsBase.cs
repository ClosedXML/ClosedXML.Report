using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FluentAssertions;
//using JetBrains.Profiler.Windows.Api;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlsxTemplateTestsBase
    {
        protected readonly ITestOutputHelper Output;

        public XlsxTemplateTestsBase(ITestOutputHelper output)
        {
            Output = output;
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
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var template = new XLTemplate(stream))
            {
                // ARRANGE
                arrangeCallback(template);

                using (var file = new MemoryStream())
                {
                    //MemoryProfiler.Dump();
                    // ACT
                    var start = DateTime.Now;
                    template.Generate();
                    Output.WriteLine(DateTime.Now.Subtract(start).ToString());
                    //MemoryProfiler.Dump();
                    template.SaveAs(file);
                    //MemoryProfiler.Dump();
                    file.Position = 0;

                    using (var wb = new XLWorkbook(file))
                    {
                        // ASSERT
                        assertCallback(wb);
                    }
                }
            }

            GC.Collect();
            //MemoryProfiler.Dump();
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
                actual.Worksheets.Count.ShouldBeEquivalentTo(expected.Worksheets.Count, $"Count of worksheets must be {expected.Worksheets.Count}");

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

                if (actualCell.Value?.ToString() != expectedCell.Value?.ToString())
                {
                    messages.Add($"Cell values are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (!string.Equals(actualCell.FormulaA1, expectedCell.FormulaA1, StringComparison.InvariantCultureIgnoreCase))
                {
                    messages.Add($"Cell formulae are not equal starting from {address}");
                    cellsAreEqual = false;
                }

                if (!expectedCell.HasFormula && actualCell.DataType != expectedCell.DataType)
                {
                    messages.Add($"Cell data types are not equal starting from {address}");
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

                for (int i = 0; i < expectedFormats.Count(); i++)
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

            if (expected.PivotTables.Count() != actual.PivotTables.Count())
                messages.Add("Pivot tables counts differ");
            else
            {
                var expPt = expected.PivotTables.ToArray();
                var actPt = actual.PivotTables.ToArray();

                for (var i = 0; i < expPt.Length; i++)
                {
                    var expectedPt = expPt[i];
                    var actualPt = actPt[i];

                    if (expectedPt.MergeAndCenterWithLabels != actualPt.MergeAndCenterWithLabels)
                        messages.Add($"Pivot table {expectedPt.Name} MergeAndCenterWithLabels are not equal");
                    if (expectedPt.ShowExpandCollapseButtons != actualPt.ShowExpandCollapseButtons)
                        messages.Add($"Pivot table {expectedPt.Name} ShowExpandCollapseButtons are not equal");
                    if (expectedPt.ClassicPivotTableLayout != actualPt.ClassicPivotTableLayout)
                        messages.Add($"Pivot table {expectedPt.Name} ClassicPivotTableLayout are not equal");
                    if (expectedPt.AutofitColumns != actualPt.AutofitColumns)
                        messages.Add($"Pivot table {expectedPt.Name} AutofitColumns are not equal");
                    if (expectedPt.SortFieldsAtoZ != actualPt.SortFieldsAtoZ)
                        messages.Add($"Pivot table {expectedPt.Name} SortFieldsAtoZ are not equal");
                    if (expectedPt.PreserveCellFormatting != actualPt.PreserveCellFormatting)
                        messages.Add($"Pivot table {expectedPt.Name} PreserveCellFormatting are not equal");
                    if (expectedPt.ShowGrandTotalsColumns != actualPt.ShowGrandTotalsColumns)
                        messages.Add($"Pivot table {expectedPt.Name} ShowGrandTotalsColumns are not equal");
                    if (expectedPt.ShowGrandTotalsRows != actualPt.ShowGrandTotalsRows)
                        messages.Add($"Pivot table {expectedPt.Name} ShowGrandTotalsRows are not equal");
                    if (expectedPt.SaveSourceData != actualPt.SaveSourceData)
                        messages.Add($"Pivot table {expectedPt.Name} SaveSourceData are not equal");
                    if (expectedPt.FilterAreaOrder != actualPt.FilterAreaOrder)
                        messages.Add($"Pivot table {expectedPt.Name} FilterAreaOrder are not equal");
                    if (expectedPt.RefreshDataOnOpen != actualPt.RefreshDataOnOpen)
                        messages.Add($"Pivot table {expectedPt.Name} RefreshDataOnOpen are not equal");
                    if (expectedPt.SourceRange.RangeAddress.ToString() != actualPt.SourceRange.RangeAddress.ToString())
                        messages.Add($"Pivot table {expectedPt.Name} SourceRange are not equal");
                    if (expectedPt.TargetCell.Address.ToString() != actualPt.TargetCell.Address.ToString())
                        messages.Add($"Pivot table {expectedPt.Name} TargetCell are not equal");
                    
                    if (expectedPt.RowLabels.Count() != actualPt.RowLabels.Count())
                        messages.Add($"Pivot table {expectedPt.Name} RowLabels counts differ");
                    else
                        PivotStylesCompare(expectedPt.RowLabels, actualPt.RowLabels, messages);

                    if (expectedPt.ColumnLabels.Count() != actualPt.ColumnLabels.Count())
                        messages.Add($"Pivot table {expectedPt.Name} ColumnLabels counts differ");
                    else
                        PivotStylesCompare(expectedPt.ColumnLabels, actualPt.ColumnLabels, messages);

                    if (expectedPt.ReportFilters.Count() != actualPt.ReportFilters.Count())
                        messages.Add($"Pivot table {expectedPt.Name} ReportFilters counts differ");
                    else
                        PivotStylesCompare(expectedPt.ReportFilters, actualPt.ReportFilters, messages);

                    if (expectedPt.Values.Count() != actualPt.Values.Count())
                        messages.Add($"Pivot table {expectedPt.Name} ReportFilters counts differ");
                    else
                    {
                        foreach (var expVal in expectedPt.Values)
                        {
                            var actRLbl = actualPt.Values.Get(expVal.SourceName);
                            if (actRLbl.SummaryFormula != expVal.SummaryFormula)
                                messages.Add($"Pivot table SummaryFormula are not equal");
                        }
                    }
                }
            }

            return !messages.Any();
        }

        private static void PivotStylesCompare(IXLPivotFields expectedFields, IXLPivotFields actualFields, IList<string> messages)
        {
            foreach (var expRLbl in expectedFields)
            {
                var actRLbl = actualFields.Get(expRLbl.SourceName);
                if (actRLbl.SubtotalCaption != expRLbl.SubtotalCaption)
                    messages.Add($"Pivot table SubtotalCaption are not equal");

                if (actRLbl.CustomName != expRLbl.CustomName)
                    messages.Add($"Pivot table CustomName are not equal");

                if (actRLbl.StyleFormats.Label.Style.ToString() != expRLbl.StyleFormats.Label.Style.ToString())
                    messages.Add($"Pivot table Label styles are not equal");

                if (actRLbl.StyleFormats.Header.Style.ToString() != expRLbl.StyleFormats.Header.Style.ToString())
                    messages.Add($"Pivot table Header styles are not equal");

                if (actRLbl.StyleFormats.Subtotal.Label.Style.ToString() != expRLbl.StyleFormats.Subtotal.Label.Style.ToString())
                    messages.Add($"Pivot table Subtotal styles are not equal");

                DataValueFormatsCompare(messages, "Subtotal", actRLbl.StyleFormats.Subtotal.DataValuesFormats, expRLbl.StyleFormats.Subtotal.DataValuesFormats);

                DataValueFormatsCompare(messages, "Data", actRLbl.StyleFormats.DataValuesFormats, expRLbl.StyleFormats.DataValuesFormats);
            }
        }

        private static void DataValueFormatsCompare(IList<string> messages, string targetName, IEnumerable<IXLPivotValueStyleFormat> actFormats,
            IEnumerable<IXLPivotValueStyleFormat> expFormats)
        {
            if (actFormats.Count() != expFormats.Count())
                messages.Add($"Pivot table {targetName} DataValuesFormats counts differ");

            var expValFormats = expFormats.ToList();
            var actValFormats = actFormats.ToList();
            for (var i = 0; i < expValFormats.Count; i++)
            {
                var expValFmt = expValFormats[i];
                var actValFmt = actValFormats[i];

                if (actValFmt.PivotField.SourceName != expValFmt.PivotField.SourceName)
                    messages.Add($"Pivot table DataValuesFormat PivotField.SourceName are not equals");
                if (actValFmt.Outline != expValFmt.Outline)
                    messages.Add($"Pivot table DataValuesFormat Outline are not equals");
                if (actValFmt.Style.ToString() != expValFmt.Style.ToString())
                    messages.Add($"Pivot table DataValuesFormat Style are not equals");
                if (actValFmt.AppliesTo != expValFmt.AppliesTo)
                    messages.Add($"Pivot table DataValuesFormat AppliesTo are not equals");
            }
        }
    }
}
