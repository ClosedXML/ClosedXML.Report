namespace ClosedXML.Report.Options
{
    public class ColsFitTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlRange = context.Range;
            var xlCell = Cell.GetXlCell(xlRange);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();
            var ws = xlRange.Worksheet;

            // whole workbook
            if (cellAddr == "A1")
            {
                foreach (var worksheet in ws.Workbook.Worksheets)
                {
                    worksheet.Columns().AdjustToContents();
                }
            }
            // whole worksheet
            else if (cellAddr == "A2")
            {
                ws.Columns().AdjustToContents();
            }
            // worksheet column
            else if (cellRow == 1)
            {
                ws.Column(cellClmn).AdjustToContents(ws.FirstRowUsed().RowNumber(), ws.LastRowUsed().RowNumber());
            }
            // whole range
            else if (IsSpecialRangeCell(xlCell))
            {
                ws.Columns(xlRange.FirstColumnUsed().ColumnNumber(), xlRange.LastColumnUsed().ColumnNumber())
                    .AdjustToContents(xlRange.FirstRowUsed().RowNumber()-1, xlRange.LastRowUsed().RowNumber());
            }
            // range column
            if (IsSpecialRangeRow(xlCell))
            {
                ws.Column(cellClmn).AdjustToContents(xlRange.FirstRowUsed().RowNumber(), xlRange.LastRowUsed().RowNumber());
            }
            // only one cell
            else
            {
                ws.Column(cellClmn).AdjustToContents(cellRow, cellRow);
            }
        }
    }
}
