using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class DeleteTagTests: TagTests
    {
        [Fact]
        public void TagInA2CellShouldDeleteWorksheet()
        {
            var tag = CreateNotInRangeTag<DeleteTag>(_ws.Cell("A2"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), new DataSource(new object[0])));

            _wb.Worksheets.Count.Should().Be(0);
        }

        [Fact]
        public void TagInFirstWorksheetRowCellShouldDeleteWholeColumn()
        {
            _ws.Cell("B5").Value = 2.0;
            _ws.Cell("C5").Value = 3.0;
            _ws.Cell("D5").Value = 4.0;
            var tag = CreateNotInRangeTag<DeleteTag>(_ws.Cell("C1"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), new DataSource(new object[0])));

            _ws.Cell("B5").Value.Should().Be(2.0);
            _ws.Cell("C5").Value.Should().Be(4.0);
        }

        [Fact]
        public void TagInFirstWorksheetColumnCellShouldDeleteWholeRow()
        {
            _ws.Cell("D3").Value = 3.0;
            _ws.Cell("D4").Value = 4.0;
            _ws.Cell("D5").Value = 5.0;
            var tag = CreateNotInRangeTag<DeleteTag>(_ws.Cell("A4"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), new DataSource(new object[0])));

            _ws.Cell("D3").Value.Should().Be(3.0);
            _ws.Cell("D4").Value.Should().Be(5.0);
        }

        [Fact]
        public void TagInDataCellRangeOptionRowShouldDeleteThisColumn()
        {
            var rng = FillData();

            var tag = CreateInRangeTag<DeleteTag>(rng, rng.Cell("B2"));
            tag.Execute(new ProcessingContext(_ws.Range("B5", "F15"), new DataSource(new object[0])));

            rng.Cell("A1").Value.Should().Be("Alice");
            rng.Cell("B1").Value.Should().Be("Wonderland");
        }

        private IXLRange FillData()
        {
            var rng = _ws.Range("B5", "F6");
            _ws.Cell("A5").Value = "Not in range";
            _ws.Cell("G5").Value = "Not in range";
            _ws.Cell("B4").InsertTable(GenerateItems());
            _ws.Columns(1, 7).Width = 3;
            return rng;
        }

        private IEnumerable<object> GenerateItems()
        {
            return new List<object>
            {
                new {FirstName = "Alice", LastName = "Liddell", Country = "Wonderland", Address = "Westminster, London, England", BirthDay = DateTime.Parse("1852-05-04").ToString("d")},
                new {FirstName = "Lewis", LastName = "Carroll", Country = "UK", Address = "Daresbury, Cheshire, England", BirthDay = DateTime.Parse("1832-01-27").ToString("d")},
                new {FirstName = "Charles", LastName = "Dodgson", Country = "UK", Address = "Daresbury, Cheshire, England", BirthDay = DateTime.Parse("1832-01-27").ToString("d")},
            };
        }
    }
}
