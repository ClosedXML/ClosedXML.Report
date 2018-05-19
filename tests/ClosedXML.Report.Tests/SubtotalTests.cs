using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class SubtotalTests: XlsxTemplateTestsBase
    {
        private IXLRange _rng;
        private XLWorkbook _workbook;

        public SubtotalTests(ITestOutputHelper output) : base(output)
        {
        }

        private void LoadTemplate(string fileTemplate)
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, fileTemplate);
            _workbook = new XLWorkbook(fileName);
            _rng = _workbook.Range("range1");
        }

        [Fact]
        public void SubtotalsTest()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 });
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false);
#if SAVE_OUTPUT
            _workbook.SaveAs("Output\\XlExtensions_Subtotals.xlsx");
#endif
            CompareWithGauge(_workbook, "XlExtensions_Subtotals.xlsx");
        }

        [Fact]
        public void SubtotalsReplaceTest()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 });
            _rng.Subtotal(3, "sum", new[] { 5, 7 });
#if SAVE_OUTPUT
            _workbook.SaveAs("Output\\XlExtensions_SubtotalsReplace.xlsx");
#endif
            CompareWithGauge(_workbook, "XlExtensions_SubtotalsReplace.xlsx");
        }

        [Fact]
        public void SummaryAbove()
        {
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 }, summaryAbove: true);
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false, summaryAbove: true);
#if SAVE_OUTPUT
            _workbook.SaveAs("Output\\Subtotal_SummaryAbove.xlsx");
#endif
            CompareWithGauge(_workbook, "Subtotal_SummaryAbove.xlsx");
        }

        [Fact]
        public void PageBreaks()
        { 
            LoadTemplate("9_plaindata.xlsx");
            _rng.Subtotal(2, "sum", new[] { 5, 7 }, pageBreaks: true);
            _rng.Subtotal(3, "sum", new[] { 5, 7 }, false, true);
#if SAVE_OUTPUT
            _workbook.SaveAs("Output\\Subtotal_PageBreaks.xlsx");
#endif
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
#if SAVE_OUTPUT
            _workbook.SaveAs("Output\\Subtotal_WithHeaders.xlsx");
#endif
            CompareWithGauge(_workbook, "Subtotal_WithHeaders.xlsx");
        }
    }
}
