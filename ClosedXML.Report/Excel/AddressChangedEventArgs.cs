using System;
using ClosedXML.Excel;

namespace ClosedXML.Report.Excel
{
    public class AddressChangedEventArgs : EventArgs
    {
        public AddressChangedEventArgs(IXLRangeBase range, IXLRangeAddress oldAddress, IXLRangeAddress newAddress)
        {
            NewAddress = newAddress;
            Range = range;
            OldAddress = oldAddress;
        }

        public IXLRangeBase Range { get; private set; }
        public IXLRangeAddress OldAddress { get; private set; }
        public IXLRangeAddress NewAddress { get; private set; }

        public int NewRowsCount
        {
            get { return NewAddress.LastAddress.RowNumber - NewAddress.FirstAddress.RowNumber + 1; }
        }

        public int NewColumnsCount
        {
            get { return NewAddress.LastAddress.ColumnNumber - NewAddress.FirstAddress.ColumnNumber + 1; }
        }

        public int GetHeightDiff()
        {
            var oldRowsCnt = OldAddress.LastAddress.RowNumber - OldAddress.FirstAddress.RowNumber + 1;
            return NewRowsCount - oldRowsCnt;
        }

        public int GetWidthDiff()
        {
            var oldColsCnt = OldAddress.LastAddress.ColumnNumber - OldAddress.FirstAddress.ColumnNumber + 1;
            return NewColumnsCount - oldColsCnt;
        }
    }
}