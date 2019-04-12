using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public abstract class OptionTag
    {
        internal byte PriorityKey { get; set; }

        public Dictionary<string, string> Parameters { get; internal set; }
        public TemplateCell Cell { get; internal set; }
        public TagsList List { get; internal set; }
        public string Name { get; internal set; }
        public bool Enabled { get; set; }
        public byte Priority { get; internal set; }

        private string _rangeOptionsRow;
        public IXLRangeAddress RangeOptionsRow
        {
            get
            {
                return _rangeOptionsRow == null
                    ? null
                    : Cell.XLCell.Worksheet.Range(_rangeOptionsRow).RangeAddress;
            }
            set { _rangeOptionsRow = value.ToString(); }
        }

        private int _column;
        public int Column
        {
            get { return _column > 0 ? _column : (_column = Cell.Column); }
            protected set { _column = value; }
        }

        protected OptionTag()
        {
            Enabled = true;
            Parameters = new Dictionary<string, string>();
        }

        private IXLRange _range;
        public IXLRange Range
        {
            get { return _range; }
            set
            {
                if (Equals(_range, value)) return;
                SetRange(value);
            }
        }

        protected virtual void SetRange(IXLRange value)
        {
            _range = value;
        }

        protected bool IsSpecialRangeCell(IXLCell cell)
        {
            var cellRow = cell.WorksheetRow().RowNumber();
            var cellClmn = cell.WorksheetColumn().ColumnNumber();
            return cellRow == RangeOptionsRow?.LastAddress.RowNumber && cellClmn == RangeOptionsRow.FirstAddress.ColumnNumber;
        }

        protected bool IsSpecialRangeRow(IXLCell cell)
        {
            return cell.Address.RowNumber == RangeOptionsRow?.LastAddress.RowNumber;
        }

        public virtual void Execute(ProcessingContext context)
        {
            Enabled = false;
        }

        public string GetParameter(string name)
        {
            return Parameters.TryGetValue(name.ToLower(), out var val) ? val : null;
        }

        public bool HasParameter(string name)
        {
            return Parameters.ContainsKey(name.ToLower());
        }

        public object Clone()
        {
            var clone = (OptionTag)MemberwiseClone();
            return clone;
        }
    }
}
