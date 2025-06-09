using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using FluentAssertions;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class TempSheetBufferTests
    {
        [Fact]
        public void NamedRangesAreRemovedWithTempSheet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");

                var tempSheetBuffer = new TempSheetBuffer(wb);
                wb.DefinedNames.Add("Temp range", tempSheetBuffer.GetRange(
                    tempSheetBuffer.GetCell(1, 1).Address,
                    tempSheetBuffer.GetCell(4, 4).Address));

                wb.DefinedNames.Count().Should().Be(1, "global named range is supposed to be added");
                tempSheetBuffer.Dispose();
                wb.DefinedNames.Count().Should().Be(0, "named range should be deleted with the temp buffer");
            }
        }

        [Fact]
        public void CanRenderRangeForEmptySet()
        {
            // If bound range has no items and the option row is empty, remove the whole area, including the option row.
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Range("A2:B3").AddToNamed("List");
                ws.Cell("B2").Value = "{{item}}";
                ws.Cell("B4").Value = "Cell below";

                var template = new XLTemplate(wb);
                template.AddVariable("List", new List<string>());
                template.Generate();

                ws.Cell("B2").GetString().Should().Be("Cell below");
                ws.Cell("B3").GetString().Should().Be("");
            }
        }

        [Fact]
        public void InnerRange()
        {
            using (var wb = new XLWorkbook())
            {
                // Arrange.
                const string rangeName = "List";
                const string innerRangeName = "CustomTotals";
                const string totalsName = "Totals";
                var list = new List<string> { "Value1", "Value2" };
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("B1").Value = "Header";
                ws.Range("A2:C3").AddToNamed(rangeName);
                ws.Cell("B2").Value = "{{index+1}}";
                ws.Cell("C2").Value = "{{item}}";
                ws.Cell("B3").Value = totalsName;
                ws.Range("C3:C3").AddToNamed(innerRangeName);

                // Act.
                var template = new XLTemplate(wb);
                template.AddVariable(rangeName, list);
                template.AddVariable(innerRangeName, list.Count);
                template.Generate();

                // Assert.
                ws.Cell("B4").GetString().Should().Be(totalsName);
                ws.Range(rangeName).RowCount().Should().Be(3);
                ws.Range(rangeName).ColumnCount().Should().Be(3);
            }
        }
    }
}
