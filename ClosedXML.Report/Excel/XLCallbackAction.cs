using System;
using ClosedXML.Excel;

namespace ClosedXML.Report.Excel
{
    public class XLCallbackAction
    {
        public XLCallbackAction(Action<IXLRangeAddress, Int32> action)
        {
            this.Action = action;
        }

        public Action<IXLRangeAddress, Int32> Action { get; set; }
    }
}