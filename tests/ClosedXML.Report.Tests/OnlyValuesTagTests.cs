using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Options;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class OnlyValuesTagTests : TagTests
    {
        [Fact]
        public void TagInA2CellShouldReplaceAllFormulasOnWorksheet()
        {
            FillData();
            var tag = CreateNotInRangeTag<OnlyValuesTag>(_ws.Cell("A2"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), new DataSource(new object[0]), new FormulaEvaluator()));

            _ws.CellsUsed(c => c.HasFormula).Should().BeEmpty();
        }

        [Fact]
        public void TagInFirstCellRangeOptionRowShouldReplaceAllFormulasInRange()
        {
            var rng = FillData();

            var tag = CreateInRangeTag<OnlyValuesTag>(rng, rng.Cell(2, 1));
            tag.Execute(new ProcessingContext(_ws.Range("B5", "F15"), new DataSource(new object[0]), new FormulaEvaluator()));

            rng.CellsUsed(c => c.HasFormula).Should().BeEmpty();
            _ws.Cell("B3").HasFormula.Should().BeTrue();
        }

        [Fact]
        public void TagInRangeCellShouldReplaceAllFormulasOnlyThisColumnInRange()
        {
            var rng = FillData();
            var dataRng = _ws.Range("B5", "D7");

            var tag = CreateInRangeTag<OnlyValuesTag>(rng, rng.Cell(1, 2));
            tag.Execute(new ProcessingContext(dataRng, new DataSource(new object[0]), new FormulaEvaluator()));

            dataRng.Column(1).Cells(c => c.HasFormula).Count().Should().Be(3);
            dataRng.Column(2).Cells(c => c.HasFormula).Should().BeEmpty();
            dataRng.Column(3).Cells(c => c.HasFormula).Count().Should().Be(3);
            _ws.Cell("B3").HasFormula.Should().BeTrue();
        }

        [Fact]
        public void TagNotInRangeCellShouldReplaceFormulasOnlyThisCell()
        {
            var rng = FillData();

            var tag = CreateNotInRangeTag<OnlyValuesTag>(_ws.Cell("B3"));
            tag.Execute(new ProcessingContext(_ws.AsRange(), null, new FormulaEvaluator()));

            _ws.Cell("B3").HasFormula.Should().BeFalse();
            rng.Cells().All(c => c.HasFormula).Should().BeTrue();
        }

        private IXLRange FillData()
        {
            var rng = _ws.Range("B5", "D6");
            _ws.Cell("B3").FormulaA1 = "99+99";
            for (int r = 1; r <= 3; r++)
            {
                for (int c = 1; c <= 3; c++)
                {
                    rng.Cell(r, c).FormulaA1 = r + "+" + c;
                }
            }
            return rng;
        }
    }
}
