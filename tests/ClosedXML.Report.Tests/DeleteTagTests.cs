using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class DeleteTagTests : TagTests
    {
        [Fact]
        public void TagInA2CellShouldDeleteWorksheet()
        {
            CreateAndExecuteTags(new[] { "A2" });

            _wb.Worksheets.Count.Should().Be(0);
        }

        [Fact]
        public void TagInFirstWorksheetRowCellShouldDeleteWholeColumn()
        {
            _ws.Cell("B5").Value = 2.0;
            _ws.Cell("C5").Value = 3.0;
            _ws.Cell("D5").Value = 4.0;
            CreateAndExecuteTags(new[] { "C1" });

            _ws.Cell("B5").GetDouble().Should().Be(2.0);
            _ws.Cell("C5").GetDouble().Should().Be(4.0);
        }

        [Fact]
        public void TagInFirstWorksheetColumnCellShouldDeleteWholeRow()
        {
            _ws.Cell("D3").Value = 3.0;
            _ws.Cell("D4").Value = 4.0;
            _ws.Cell("D5").Value = 5.0;
            CreateAndExecuteTags(new[] { "A4" });

            _ws.Cell("D3").GetDouble().Should().Be(3.0);
            _ws.Cell("D4").GetDouble().Should().Be(5.0);
        }

        [Fact]
        public void TagsShouldDeleteFromLastToFirstCell()
        {
            _ws.Cell("B5").Value = 2.0;
            _ws.Cell("C5").Value = 3.0;
            _ws.Cell("D5").Value = 4.0;
            _ws.Cell("E5").Value = 5.0;
            CreateAndExecuteTags(new[] { "A3", "A4", "C1", "D1" });

            _ws.Cell("B3").GetDouble().Should().Be(2.0);
            _ws.Cell("C3").GetDouble().Should().Be(5.0);
        }

        [Fact]
        public void TagInDataCellRangeOptionRowShouldDeleteThisColumn()
        {
            var rng = FillData();

            CreateAndExecuteTags(new[] { "B2" }, isInRange: true, range: rng);

            rng.Cell("A1").GetText().Should().Be("Alice");
            rng.Cell("B1").GetText().Should().Be("Wonderland");
        }

        private void CreateAndExecuteTags(IEnumerable<string> cells, bool isInRange = false, IXLRange range = null)
        {
            var errorsList = new TemplateErrors();
            var tagList = new TagsList(errorsList);
            foreach (var cell in cells)
            {
                tagList.Add(isInRange
                    ? CreateInRangeTag<DeleteTag>(range, range.Cell(cell))
                    : CreateNotInRangeTag<DeleteTag>(_ws.Cell(cell)));
            }

            tagList.Execute(new ProcessingContext(isInRange ? range : _ws.AsRange(), new DataSource(Array.Empty<object>()), new FormulaEvaluator()));
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
