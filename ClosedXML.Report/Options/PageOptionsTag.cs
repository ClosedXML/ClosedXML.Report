using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class PageOptionsTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var ws = Range.Worksheet;

            // workbook
            if (cellAddr == "A1")
            {
                if (HasParameter("Wide")) ws.Workbook.PageOptions.PagesWide = GetParameter("Wide").AsInt(1);
                if (HasParameter("Tall")) ws.Workbook.PageOptions.PagesTall = GetParameter("Tall").AsInt(1);
                if (HasParameter("Landscape")) ws.Workbook.PageOptions.PageOrientation = XLPageOrientation.Landscape;
            }
            // worksheet
            else if (cellAddr == "A2")
            {
                if (HasParameter("Wide")) ws.PageSetup.PagesWide = GetParameter("Wide").AsInt(1);
                if (HasParameter("Tall")) ws.PageSetup.PagesTall = GetParameter("Tall").AsInt(1);
                if (HasParameter("Landscape")) ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            }
        }

        public override byte Priority { get { return 0; } }
    }
}