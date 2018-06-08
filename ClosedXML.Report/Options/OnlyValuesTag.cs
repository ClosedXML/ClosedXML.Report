/*
================================================
OPTION          OBJECTS   
================================================
"OnlyValues"    Worksheet, Range, Range Column, Cell
================================================
*/

using ClosedXML.Excel;
using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public class OnlyValuesTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var range = context.Range;
            var xlCell = Cell.GetXlCell(range);
            var cellAddr = xlCell.Address.ToStringRelative(false);

            // whole worksheet or range
            if (IsSpecialRangeCell(xlCell) || cellAddr == "A2")
            {
                range.CellsUsed(i => i.HasFormula)
                    .ForEach(c => c.Value = c.Value);
            }
            // range column
            else if (RangeOptionsRow != null)
            {
                var addrInRange = xlCell.Relative(Range.RangeAddress.FirstAddress);
                context.Range.Column(addrInRange.ColumnNumber)
                    .CellsUsed(i => i.HasFormula)
                    .ForEach(c => c.Value = c.Value);
            }
            // one cell
            else if (Cell.CellType == TemplateCellType.Formula)
            {
                xlCell.Value = xlCell.Value;
            }
        }
        public override byte Priority => 40;
    }
}
