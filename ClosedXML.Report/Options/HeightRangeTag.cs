using System.Linq;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class HeightRangeTag : RangeOptionTag
    {
        public int Height => Parameters.Any() ? Parameters.First().Key.AsInt() : 0;

        public override void Execute(ProcessingContext context)
        {
            var firstRow = context.Range.FirstRowUsed();
            var lastRow = context.Range.LastRowUsed();

            Range.Worksheet.Rows(firstRow.WorksheetRow().RowNumber(), lastRow.WorksheetRow().RowNumber()).Height = Height;
        }
    }
}
