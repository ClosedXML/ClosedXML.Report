using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public class OnlyValuesTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var range = context != null ? context.Range : Range;
            if (Cell.CellType == TemplateCellType.Formula)
            {
                var xlCell = Cell.GetXlCell(range);
                xlCell.Value = xlCell.Value;
            }
            else
            {
                range.CellsUsed(i => i.HasFormula)
                        .ForEach(c => c.Value = c.Value);
            }
        }
        public override byte Priority => 40;
    }
}
