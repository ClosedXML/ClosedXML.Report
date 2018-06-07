using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Report.Options
{
    public class AutoFilterTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellRow = xlCell.WorksheetRow().RowNumber();
            var cellClmn = xlCell.WorksheetColumn().ColumnNumber();

            var itemsCnt = context.Value is DataSource ds ? ds.GetAll().Length : 0;
            if (cellRow == context.Range.RangeAddress.LastAddress.RowNumber - itemsCnt + 1 && cellClmn == 1)
            {
                context.Range.Range(context.Range.FirstCell().CellRight(), context.Range.LastCell()).FirstRow().RowAbove().SetAutoFilter();
            }
        }

        public override byte Priority => 10;
    }
}
