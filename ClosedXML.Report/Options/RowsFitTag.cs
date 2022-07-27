namespace ClosedXML.Report.Options
{
    public class RowsFitTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlRange = context.Range;
            var xlCell = Cell.GetXlCell(xlRange);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();
            var ws = Range.Worksheet;

            // whole workbook
            if (cellAddr == "A1")
            {
                foreach (var worksheet in ws.Workbook.Worksheets)
                {
                    worksheet.Rows().AdjustToContents();
                }
            }
            // whole worksheet
            else if (cellAddr == "A2")
            {
                ws.Rows().AdjustToContents();
            }
            // worksheet row
            else if (cellClmn == 1)
            {
                var firstColumnUsed = ws.FirstColumnUsed();
                if (firstColumnUsed != null)
                    ws.Rows(cellRow, cellRow).AdjustToContents(firstColumnUsed.ColumnNumber(), ws.LastColumnUsed().ColumnNumber());
            }
            // whole range
            else if (IsSpecialRangeCell(xlCell))
            {
                var firstUsed = xlRange.FirstCellUsed();
                var lastUsed = xlRange.LastCellUsed();

                if (firstUsed != null && lastUsed != null)
                {
                    ws.Rows(firstUsed.WorksheetRow().RowNumber(), lastUsed.WorksheetRow().RowNumber())
                        .AdjustToContents(firstUsed.WorksheetColumn().ColumnNumber(), lastUsed.WorksheetColumn().ColumnNumber());
                }
            }
            // range row
            if (cellClmn == xlRange.RangeAddress.FirstAddress.ColumnNumber)
            {
                var firstColumnUsed = xlRange.FirstColumnUsed();
                if (firstColumnUsed != null)
                    ws.Rows(cellRow, cellRow).AdjustToContents(firstColumnUsed.ColumnNumber(), xlRange.LastColumnUsed().ColumnNumber());
            }
            // only one cell
            else
            {
                ws.Rows(cellRow, cellRow).AdjustToContents(cellClmn, cellClmn);
            }
        }
    }
}
