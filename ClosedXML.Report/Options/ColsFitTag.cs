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
                var firstRowUsed = ws.FirstRowUsed();
                if (firstRowUsed != null)
                    ws.Column(cellClmn).AdjustToContents(firstRowUsed.RowNumber(), ws.LastRowUsed().RowNumber());
            }
            // whole range
            else if (IsSpecialRangeCell(xlCell))
            {
                var firstUsed = xlRange.FirstCellUsed();
                var lastUsed = xlRange.LastCellUsed();

                if (firstUsed != null && lastUsed != null)
                {
                    ws.Columns(firstUsed.WorksheetColumn().ColumnNumber(), lastUsed.WorksheetColumn().ColumnNumber())
                        .AdjustToContents(firstUsed.WorksheetRow().RowNumber() - 1, lastUsed.WorksheetRow().RowNumber());
                }
            }
            // range column
            else if (IsSpecialRangeRow(xlCell))
            {
                var firstRowUsed = xlRange.FirstRowUsed();
                if (firstRowUsed != null)
                    ws.Column(cellClmn).AdjustToContents(firstRowUsed.RowNumber(), xlRange.LastRowUsed().RowNumber());
            }
            // only one cell
            else
            {
                ws.Column(cellClmn).AdjustToContents(cellRow, cellRow);
            }
        }
    }
}
