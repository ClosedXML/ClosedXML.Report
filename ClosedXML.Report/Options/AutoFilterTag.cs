using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public class AutoFilterTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            if (IsSpecialRangeCell(xlCell))
            {
                context.Range.Range(context.Range.FirstCell().CellRight(), context.Range.LastCell())
                    .FirstRow()
                    .RowAbove()
                    .SetAutoFilter();
            }
        }
    }
}
