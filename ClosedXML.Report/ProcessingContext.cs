using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public class ProcessingContext
    {
        public ProcessingContext(IXLRange range, object value)
        {
            Range = range;
            Value = value;
        }

        public object Value { get; set; }
        public IXLRange Range { get; set; }
    }
}