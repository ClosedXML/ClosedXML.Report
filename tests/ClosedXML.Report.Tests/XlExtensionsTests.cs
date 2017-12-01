using System.IO;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXML.Report.Tests
{
    public class XlExtensionsTests: XlsxTemplateTestsBase
    {
        public XlExtensionsTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void OffsetTest()
        {
            var fileName = Path.Combine(TestConstants.TemplatesFolder, "1.xlsx");
            using (var workbook = new XLWorkbook(fileName))
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
            using (var workbook = new XLWorkbook(fileName))
            {
                var pars = workbook.Worksheet(1).GetRangeParameters("SUBTOTALS(9,I36:I38, U32");
                pars.Length.Should().Be(2);
                pars[0].Key.Should().Be("I36:I38");
                pars[0].Value.ToStringRelative().Should().Be("I36:I38");
                pars[1].Key.Should().Be("U32");
                pars[1].Value.ToStringRelative().Should().Be("U32:U32");
            }
        }
    }
}