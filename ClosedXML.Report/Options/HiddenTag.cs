namespace ClosedXML.Report.Options
{
    public class HiddenTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var range = context != null ? context.Range : Range;
            var xlCell = Cell.GetXlCell(range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();
            var ws = range.Worksheet;

            // worksheet
            if (cellAddr == "A2")
            {
                ws.Hide();
            }
            // whole range
            else if (cellRow == range.LastRow().RowNumber() && cellClmn == 1)
            {
                range.Worksheet.Hide();
            }
        }
        public override byte Priority { get { return 0; } }
    }
}