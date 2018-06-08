namespace ClosedXML.Report.Options
{
    public class ColsFitTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlRange = context?.Range ?? Range;
            var xlCell = Cell.GetXlCell(xlRange);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();
            var ws = xlRange.Worksheet;
            var itemsCnt = context?.Value is DataSource ds ? ds.GetAll().Length : 0;

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
                ws.Columns(cellClmn, cellClmn).AdjustToContents(ws.FirstRowUsed().RowNumber(), ws.LastRowUsed().RowNumber());
            }
            // whole range
            else if (cellRow == Range.RangeAddress.LastAddress.RowNumber && cellClmn == 1)
            {
                ws.Columns(xlRange.FirstColumnUsed().ColumnNumber(), xlRange.LastColumnUsed().ColumnNumber())
                    .AdjustToContents(xlRange.FirstRowUsed().RowNumber()-1, xlRange.LastRowUsed().RowNumber());
            }
            // range column
            if (cellRow == xlRange.RangeAddress.FirstAddress.RowNumber)
            {
                ws.Columns(cellClmn, cellClmn).AdjustToContents(xlRange.FirstRowUsed().RowNumber(), xlRange.LastRowUsed().RowNumber());
            }
            // only one cell
            else
            {
                ws.Columns(cellClmn, cellClmn).AdjustToContents(cellRow, cellRow);
            }
        }
        public override byte Priority { get { return 0; } }
    }
}
