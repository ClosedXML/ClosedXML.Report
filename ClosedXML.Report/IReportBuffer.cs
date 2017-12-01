using System;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public interface IReportBuffer: IDisposable
    {
        IXLRange CopyTo(IXLRange range);
        void Clear();
    }
}