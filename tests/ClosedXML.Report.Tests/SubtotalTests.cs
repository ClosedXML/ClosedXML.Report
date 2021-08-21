using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class SubtotalTests: XlsxTemplateTestsBase, IDisposable
    {
        private IXLRange _rng;
        private XLWorkbook _workbook;
        private FileStream _stream;

        public SubtotalTests(ITestOutputHelper output) : base(output)
        {
        }

        private void LoadTemplate(string fileTemplate)
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, fileTemplate);
            _stream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _workbook = new XLWorkbook(_stream);
            _rng = _workbook.Range("range1");
        }

        [Fact]
        public void SubtotalsTest()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 });
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false);
            CompareWithGauge(_workbook, "XlExtensions_Subtotals.xlsx");
        }

        [Fact]
        public void SubtotalsReplaceTest()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 });
            _rng.Subtotal(3, "sum", new[] { 5, 7 });
            CompareWithGauge(_workbook, "XlExtensions_SubtotalsReplace.xlsx");
        }

        [Fact]
        public void SummaryAbove()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 }, summaryAbove: true);
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false, summaryAbove: true);
            CompareWithGauge(_workbook, "Subtotal_SummaryAbove.xlsx");
        }

        [Fact]
        public void ScanForGroupsTest()
        {
            void AssertGroup(SubtotalGroup group, string expectedAddress, string expectedTitle)
            {
                group.Range.RangeAddress.ToString().Should().Be(expectedAddress);
                group.Level.Should().Be(1);
                group.GroupTitle.Should().Be(expectedTitle);
            }

            _workbook = new XLWorkbook();
            var sheet = _workbook.AddWorksheet("test");
            sheet.Range("A2:A5").Value = "val1";
            sheet.Range("A6:A9").Value = "val2";
            sheet.Range("A10:A13").Value = "val3";

            sheet.Range("B2:B4").Value = "val4";
            sheet.Range("B5:B11").Value = "val5";
            sheet.Range("B12:B13").Value = "val6";

            sheet.Range("C2:C3").Value = "val7";
            sheet.Range("C4:C6").Value = "val8";
            sheet.Range("C7:C8").Value = "val9";
            sheet.Range("C9:C12").Value = "val7";
            sheet.Range("C13:C13").Value = "val8";

            _rng = sheet.Range("A2:C13");

            using (var subtotal = new Subtotal(_rng))
            {
                var groups = subtotal.ScanForGroups(1);
                groups.Length.Should().Be(3);
                AssertGroup(groups[0], "A2:C5", "val1");
                AssertGroup(groups[1], "A6:C9", "val2");
                AssertGroup(groups[2], "A10:C13", "val3");

                groups = subtotal.ScanForGroups(2);
                groups.Length.Should().Be(5);
                AssertGroup(groups[0], "A2:C4", "val4");
                AssertGroup(groups[1], "A5:C5", "val5");
                AssertGroup(groups[2], "A6:C9", "val5");
                AssertGroup(groups[3], "A10:C11", "val5");
                AssertGroup(groups[4], "A12:C13", "val6");

                groups = subtotal.ScanForGroups(3);
                groups.Length.Should().Be(9);
                AssertGroup(groups[0], "A2:C3", "val7");
                AssertGroup(groups[1], "A4:C4", "val8");
                AssertGroup(groups[2], "A5:C5", "val8");
                AssertGroup(groups[3], "A6:C6", "val8");
                AssertGroup(groups[4], "A7:C8", "val9");
                AssertGroup(groups[5], "A9:C9", "val7");
                AssertGroup(groups[6], "A10:C11", "val7");
                AssertGroup(groups[7], "A12:C12", "val7");
                AssertGroup(groups[8], "A13:C13", "val8");
            }
        }

        [Fact]
        public void PageBreaks()
        { 
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 }, pageBreaks: true);
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false, true);
            CompareWithGauge(_workbook, "Subtotal_PageBreaks.xlsx");
        }
        
        [Fact]
        public void WithHeaders()
        {
            LoadTemplate("91_plaindata.xlsx");
            using (var subtotal = new Subtotal(_rng))
            {
                var summaries = new[] { new SummaryFuncTag { Name = "sum", Cell=new TemplateCell { Column = 7 } } };//{new SubtotalSummaryFunc("sum", 7), };
                subtotal.AddGrandTotal(summaries);
                subtotal.GroupBy(2, summaries, true);
                subtotal.GroupBy(3, summaries, true);
                subtotal.AddHeaders(2);
                subtotal.AddHeaders(3);
            }
            CompareWithGauge(_workbook, "Subtotal_WithHeaders.xlsx");
        }

        public void Dispose()
        {
            _workbook?.Dispose();
            _stream?.Dispose();
        }
    }
}
