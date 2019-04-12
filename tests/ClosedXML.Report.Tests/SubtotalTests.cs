using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
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
            LoadTemplate("9_plaindata.xlsx");

            SubtotalGroup[] groups;
            using (var subtotal = new Subtotal(_rng))
                groups = subtotal.ScanForGroups(2);

            groups.Length.Should().Be(3);
            groups[0].Range.RangeAddress.ToString().Should().Be("C3:I26");
            groups[0].Level.Should().Be(1);
            groups[0].GroupTitle.Should().Be("Central");
            groups[1].Range.RangeAddress.ToString().Should().Be("C27:I38");
            groups[1].Level.Should().Be(1);
            groups[1].GroupTitle.Should().Be("East");
            groups[2].Range.RangeAddress.ToString().Should().Be("C39:I44");
            groups[2].Level.Should().Be(1);
            groups[2].GroupTitle.Should().Be("West");
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
                var summaries = new [] {new SubtotalSummaryFunc("sum", 7), };
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
