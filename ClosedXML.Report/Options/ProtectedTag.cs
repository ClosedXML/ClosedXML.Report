using System;

namespace ClosedXML.Report.Options
{
    public class ProtectedTag: OptionTag
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
                ws.Workbook.Protect();
            }
            // whole worksheet
            else if (cellAddr == "A2")
            {
                var passw = GetParameter("Password");
                if (string.IsNullOrEmpty(passw))
                    passw = Guid.NewGuid().ToString();
                ws.Protect(passw);
            }
            // whole range
            else if (cellRow == Range.LastRow().RowNumber() && cellClmn == 1)
            {
                Range.Style.Protection.Locked = true;
            }
            else
            {
                xlCell.Style.Protection.Locked = true;
            }
        }
        public override byte Priority { get { return 0; } }
    }
}