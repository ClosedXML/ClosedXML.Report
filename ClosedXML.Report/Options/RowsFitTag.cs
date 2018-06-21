namespace ClosedXML.Report.Options
{
    public class RowsFitTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
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
                ws.Rows(cellRow, cellRow).AdjustToContents(ws.FirstColumnUsed().ColumnNumber(), ws.LastColumnUsed().ColumnNumber());
            }
            // whole range
            else if (IsSpecialRangeCell(xlCell))
            {
                ws.Rows(context.Range.FirstRowUsed().RowNumber(), context.Range.LastRowUsed().RowNumber())
                    .AdjustToContents(context.Range.FirstColumnUsed().ColumnNumber(), context.Range.LastColumnUsed().ColumnNumber());
            }
            // range row
            if (cellClmn == context.Range.RangeAddress.FirstAddress.ColumnNumber)
            {
                ws.Rows(cellRow, cellRow).AdjustToContents(context.Range.FirstColumnUsed().ColumnNumber(), context.Range.LastColumnUsed().ColumnNumber());
            }
            // only one cell
            else
            {
                ws.Rows(cellRow, cellRow).AdjustToContents(cellClmn, cellClmn);
            }
        }
    }
}
