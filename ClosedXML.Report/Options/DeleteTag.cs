/*
================================================
OPTION          OBJECTS   
================================================
"DELETE"        Worksheet, Worksheet Row/Column, Range Column
================================================
*/

using ClosedXML.Excel;
using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public class DeleteTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var ws = Range.Worksheet;

            // whole worksheet
            if (cellAddr == "A1" || cellAddr == "A2")
            {
                ws.Workbook.Worksheets.Delete(ws.Name);
            }
            // whole column
            else if (xlCell.Address.RowNumber == 1)
            {
                ws.Column(xlCell.Address.ColumnNumber).Delete();
            }
            // whole row
            else if (xlCell.Address.ColumnNumber == 1)
            {
                ws.Row(xlCell.Address.RowNumber).Delete();
            }
            // range column
            else if (IsSpecialRangeRow(xlCell))
            {
                var addrInRange = xlCell.Relative(Range.RangeAddress.FirstAddress);
                context.Range.Column(addrInRange.ColumnNumber).Delete(XLShiftDeletedCells.ShiftCellsLeft);
            }
        }
        public override byte Priority => 0;
    }
}
