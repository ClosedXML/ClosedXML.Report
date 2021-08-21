using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public static class TagExtensions
    {
        public static bool HasTag(string value)
        {
            return value.StartsWith("<<") || value.EndsWith(">>");
        }

        public static bool IsOptionsRowEmpty(this IXLRange range)
        {
            return !range.LastRow().CellsUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges).Any();
        }
    }
}
