using System.Collections.Generic;
using ClosedXML.Excel;

namespace ClosedXML.Report.Options
{
    public abstract class OptionTag
    {
        private IXLRange _range;

        internal byte PriorityKey { get; set; }

        public Dictionary<string, string> Parameters { get; set; }
        public TemplateCell Cell { get; set; }
        public TagsList List { get; set; }
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public abstract byte Priority { get; }

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