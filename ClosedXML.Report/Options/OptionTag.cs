using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;

namespace ClosedXML.Report.Options
{
    public abstract class OptionTag
    {
        internal byte PriorityKey { get; set; }

        public Dictionary<string, string> Parameters { get; set; }
        public TemplateCell Cell { get; set; }
        public TagsList List { get; set; }
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public abstract byte Priority { get; }

        private string _rangeOptionsRow;
        public IXLRangeAddress RangeOptionsRow
        {
            get { return Cell.XLCell.Worksheet.Range(_rangeOptionsRow).Unsubscribed().RangeAddress; }
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
            string val;
            if (Parameters.TryGetValue(name.ToLower(), out val))
                return val;
            else return null;
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
