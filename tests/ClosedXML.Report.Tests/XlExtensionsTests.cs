using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlExtensionsTests : XlsxTemplateTestsBase
    {
        public XlExtensionsTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void OffsetTest()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            using (var workbook = XLWorkbook.OpenFromTemplate(fileName))
            {
                var ws = workbook.Worksheet(1);
                ws.Range("A1:B2").Offset(10, 0).RangeAddress.ToStringRelative().Should().Be("A11:B12");
                ws.Range("A1:B2").Offset(0, 10).RangeAddress.ToStringRelative().Should().Be("K1:L2");

                ws.Range("A1:B2").Offset(0, 0, 10, 10).RangeAddress.ToStringRelative().Should().Be("A1:J10");
                ws.Range("A1:B2").Offset(2, 2, 10, 10).RangeAddress.ToStringRelative().Should().Be("C3:L12");

                ws.Range("C3:D4").Offset(-2, -2).RangeAddress.ToStringRelative().Should().Be("A1:B2");
                ws.Range("C3:D4").Offset(0, 0, 2, 2).RangeAddress.ToStringRelative().Should().Be("C3:D4");
            }
        }

        [Fact]
        public void GetRangeParameters_should_return_range_address_from_formula()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            using (var workbook = XLWorkbook.OpenFromTemplate(fileName))
            {
                var pars = workbook.Worksheet(1).GetRangeParameters("SUBTOTALS(9,I36:I38, U32");
                pars.Length.Should().Be(2);
                pars[0].Key.Should().Be("I36:I38");
                pars[0].Value.ToStringRelative().Should().Be("I36:I38");
                pars[1].Key.Should().Be("U32");
                pars[1].Value.ToStringRelative().Should().Be("U32:U32");
            }
        }

        [Fact]
        public void Intersection()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                XlExtensions.Intersection(ws.Range("B9:I11"), ws.Range("D4:G16")).RangeAddress.ToString().Should().Be("D9:G11");
                XlExtensions.Intersection(ws.Range("E9:I11"), ws.Range("D4:G16")).RangeAddress.ToString().Should().Be("E9:G11");
                XlExtensions.Intersection(ws.Cell("E9").AsRange(), ws.Range("D4:G16")).RangeAddress.ToString().Should().Be("E9:E9");
                XlExtensions.Intersection(ws.Range("D4:G16"), ws.Cell("E9").AsRange()).RangeAddress.ToString().Should().Be("E9:E9");

                XlExtensions.Intersection(ws.Range("A1:C3"), ws.Range("G7:I10")).Should().BeNull();
                XlExtensions.Intersection(ws.Cell("A1").AsRange(), ws.Cell("C3").AsRange()).Should().BeNull();
                XlExtensions.Intersection(ws.Range("A1:C3"), null).Should().BeNull();
            }
        }
    }
}
