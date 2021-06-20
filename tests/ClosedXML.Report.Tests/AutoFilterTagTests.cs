using ClosedXML.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class AutoFilterTagTests: TagTests
    {
        [Fact]
        public void TagInRangeOptionRowShouldAddFiltersToRangeHeader()
        {
            var rng = _ws.Range("B5", "I6");
            var tag = CreateInRangeTag<AutoFilterTag>(rng, _ws.Cell("B6"));

            tag.Execute(new ProcessingContext(_ws.Range("B5", "I15"), new DataSource(new object[0]), new FormulaEvaluator()));

            var headerRng = rng.Range(rng.FirstCell().CellRight(), rng.LastCell())
                .FirstRow().RowAbove();
            _ws.AutoFilter.Range.Should().NotBeNull("No AutoFilter specified.");
            _ws.AutoFilter.Range.RangeAddress.Should().Be(headerRng.RangeAddress, "AutoFilter range has wrong.");
        }

        [Fact]
        public void TagNotInRangeOptionRowShouldNotAddFilters()
        {
            var tag = CreateNotInRangeTag<AutoFilterTag>(_ws.Cell("B6"));

            tag.Execute(new ProcessingContext(_ws.AsRange(), null, new FormulaEvaluator()));

            _ws.AutoFilter.Range.Should().BeNull("AutoFilter is specified.");
        }
    }
}
