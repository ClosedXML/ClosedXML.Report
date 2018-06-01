using System;
using ClosedXML.Excel;

namespace ClosedXML.Report.Excel
{
    public class SubtotalGroup
    {
        public int Level { get; private set; }
        public int Column { get; set; }
        public string GroupTitle { get; private set; }
        public IXLRange Range { get; internal set; }
        public IXLRangeRow SummaryRow { get; internal set; }
        public bool PageBreaks { get; private set; }
        public IXLRangeRow HeaderRow { get; internal set; }

        public SubtotalGroup(int level, int column, string groupTitle, IXLRange range, IXLRangeRow summaryRow, bool pageBreaks)
        {
            Column = column;
            SummaryRow = summaryRow;
            PageBreaks = pageBreaks;
            Level = level;
            GroupTitle = groupTitle;
            Range = range;
        }
    }
}
