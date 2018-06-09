using System;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;

namespace ClosedXML.Report.Tests
{
    public abstract class TagTests: IDisposable
    {
        protected XLWorkbook _wb;
        protected IXLWorksheet _ws;

        protected TagTests()
        {
            _wb = new XLWorkbook();
            _ws = _wb.AddWorksheet("Sheet1");
        }

        protected T CreateInRangeTag<T>(IXLRange rng, IXLCell cell) where T : OptionTag, new()
        {
            var relAddr = cell.Relative(rng.RangeAddress.FirstAddress);
            var tag = new T()
            {
                Cell = new TemplateCell(relAddr.RowNumber, relAddr.ColumnNumber, cell),
                Range = rng,
                RangeOptionsRow = rng.LastRow().RangeAddress,
            };
            return tag;
        }

        protected T CreateNotInRangeTag<T>(IXLCell cell) where T : OptionTag, new()
        {
            var tag = new T()
            {
                Cell = new TemplateCell(cell.Address.RowNumber, cell.Address.ColumnNumber, cell),
                Range = _ws.AsRange(),
            };
            return tag;
        }

        public void Dispose()
        {
            _wb?.Dispose();
        }
    }
}
