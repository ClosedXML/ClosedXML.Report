using System.Linq;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options;

public class HeightTag : OptionTag
{
    public int Height => Parameters.Any() ? Parameters.First().Key.AsInt() : 0;

    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);
        var cellRow = xlCell.WorksheetRow().RowNumber();

        Range.Worksheet.Row(cellRow).Height = Height;
    }
}
