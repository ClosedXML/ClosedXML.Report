using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class ColsFitTagTests: TagTests
    {
        [Fact]
        public void TagInA2CellShouldFitAllColumnsInWorksheet()
        {
            FillData();

            var tag = CreateNotInRangeTag<ColsFitTag>(_ws.Cell("A2"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), new DataSource(new object[0])));

            _ws.Column(1).Width.Should().BeGreaterThan(3);
            _ws.Column(2).Width.Should().BeGreaterThan(3);
            _ws.Column(3).Width.Should().BeGreaterThan(3);
            _ws.Column(4).Width.Should().BeGreaterThan(3);
            _ws.Column(5).Width.Should().BeGreaterThan(3);
            _ws.Column(6).Width.Should().BeGreaterThan(3);
            _ws.Column(7).Width.Should().BeGreaterThan(3);
        }

        [Fact]
        public void TagInFirstCellRangeOptionRowShouldFitAllColumnsInRange()
        {
            var rng = FillData();

            var tag = CreateInRangeTag<ColsFitTag>(rng, rng.Cell(2, 1));
            tag.Execute(new ProcessingContext(_ws.Range("B5", "F15"), new DataSource(new object[0])));

            _ws.Column(1).Width.Should().Be(3);
            _ws.Column(2).Width.Should().BeGreaterThan(3);
            _ws.Column(3).Width.Should().BeGreaterThan(3);
            _ws.Column(4).Width.Should().BeGreaterThan(3);
            _ws.Column(5).Width.Should().BeGreaterThan(3);
            _ws.Column(6).Width.Should().BeGreaterThan(3);
            _ws.Column(7).Width.Should().Be(3);
        }

        [Fact]
        public void TagInDataCellRangeOptionRowShouldFitOnlyThisColumn()
        {
            var rng = FillData();

            var tag = CreateInRangeTag<ColsFitTag>(rng, rng.Cell(2, 3));
            tag.Execute(new ProcessingContext(_ws.Range("B5", "F15"), new DataSource(new object[0])));

            _ws.Column(1).Width.Should().Be(3);
            _ws.Column(2).Width.Should().Be(3);
            _ws.Column(3).Width.Should().Be(3);
            _ws.Column(4).Width.Should().BeGreaterThan(3);
            _ws.Column(5).Width.Should().Be(3);
            _ws.Column(6).Width.Should().Be(3);
            _ws.Column(7).Width.Should().Be(3);
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
