namespace ClosedXML.Report.Options
{
    public class AutoFilterTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();

            if (cellRow == context.Range.LastRow().RowNumber() && cellClmn == 1)
            {
                context.Range.FirstRow().RowAbove().SetAutoFilter();
            }
        }

        public override byte Priority { get { return 0; } }
    }
}