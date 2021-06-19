using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public class ProcessingContext
    {
        public ProcessingContext(IXLRange range, object value, FormulaEvaluator evaluator)
        {
            Range = range;
            Value = value;
            Evaluator = evaluator;
        }

        public FormulaEvaluator Evaluator { get; private set; }
        public object Value { get; private set; }
        public IXLRange Range { get; private set; }
    }

    public class SummaryProcessingContext : ProcessingContext
    {
        public IXLRangeRow SummaryRow { get; }

        public SummaryProcessingContext(IXLRange range, IXLRangeRow summaryRow, IDataSource value, FormulaEvaluator evaluator)
            : base(range, value, evaluator)
        {
            SummaryRow = summaryRow;
        }
    }
}
