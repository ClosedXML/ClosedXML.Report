/*
Protected Option
================================================
OPTION          PARAMS                OBJECTS   
================================================
"Protected"      "Password="          Workbook, Worksheet, Range, Cell, Range Column
================================================
*/
using System;
using System.Collections;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public class ProtectedTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var xlCell = Cell.GetXlCell(context.Range);
            var cellAddr = xlCell.Address.ToStringRelative(false);
            var xlAddress = xlCell.Relative(Range.RangeAddress.FirstAddress);
            var ws = Range.Worksheet;
            
            // whole workbook
            if (cellAddr == "A1")
            {
                ws.Workbook.Protect();
            }
            // whole worksheet
            else if (cellAddr == "A2")
            {
                ProtectSheet(ws);
            }
            // whole range
            else if (xlCell.Address.RowNumber == Range.RangeAddress.LastAddress.RowNumber && xlAddress.ColumnNumber == 1)
            {
                ws.Cells().ForEach(c => { c.Style.Protection.Locked = false; });
                context.Range.Cells().ForEach(c => { c.Style.Protection.Locked = true; });
                ProtectSheet(ws);
            }
            else
            {
                ws.Cells().ForEach(c => { c.Style.Protection.Locked = false; });

                if (context.Value is DataSource)
                {
                    context.Range.Column(xlAddress.ColumnNumber).Cells()
                        .ForEach(c => { c.Style.Protection.Locked = true; });
                }
                else
                {
                    xlCell.Style.Protection.Locked = true;
                }

                ProtectSheet(ws);
            }
        }

        private void ProtectSheet(IXLWorksheet ws)
        {
            var passw = GetParameter("Password");
            if (string.IsNullOrEmpty(passw))
                passw = Guid.NewGuid().ToString();
            ws.Protect(passw);
        }

        public override byte Priority { get { return 0; } }
    }
}
