﻿using ClosedXML.Excel;
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
                wb.NamedRanges.Add("Temp range", tempSheetBuffer.GetRange(
                    tempSheetBuffer.GetCell(1, 1).Address,
                    tempSheetBuffer.GetCell(4, 4).Address));

                wb.NamedRanges.Count().Should().Be(1, "global named range is supposed to be added");
                tempSheetBuffer.Dispose();
                wb.NamedRanges.Count().Should().Be(0, "named range should be deleted with the temp buffer");
            }
        }

        [Fact]
        public void CanRenderRangeForEmptySet()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Range("A2:B3").AddToNamed("List");
                ws.Cell("B2").Value = "{{item}}";
                ws.Cell("B4").Value = "Cell below";

                var template = new XLTemplate(wb);
                template.AddVariable("List", new List<string>());
                template.Generate();

                ws.Cell("B3").Value.Should().Be("Cell below");
                ws.Cell("B4").Value.Should().Be("");
            }
        }
    }
}
