using System;
using System.Linq;
using System.Linq.Expressions;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ClosedXML.Report.Options
{
    public class SummaryFuncTag : OptionTag
    {
        internal DataSource DataSource { get; set; }

        public override void Execute(ProcessingContext context)
        {
            var summ = GetFunc(context);
            IXLRangeRow summRow;
            IXLRange calculatedRange;
            // If SummaryRow is not passed, then the last row of the passed range is accepted as the SummaryRow.
            if (context is SummaryProcessingContext summaryContext && summaryContext.SummaryRow != null)
            {
                if (context.Range.RowCount() < 1)
                    return;

                summRow = summaryContext.SummaryRow;
                calculatedRange = context.Range.Offset(0, summ.Column - 1, context.Range.RowCount(), 1);
            }
            else
            {
                if (context.Range.RowCount() < 2)
                    return;

                summRow = context.Range.LastRow();
                calculatedRange = context.Range.Offset(0, summ.Column - 1, context.Range.RowCount() - 1, 1);
            }
            var items = summ.DataSource.GetAll();

            if (items == null || items.Length == 0)
            {
                if (summ.DefaultValueForEmptySource != null)
                {
                    summRow.Cell(summ.Column).Value = XLCellValueConverter.FromObject(summ.DefaultValueForEmptySource);
                }
            }
            else if (summ.FuncNum == 0)
            {
                var value = summ.Calculate((IDataSource)context.Value);
                summRow.Cell(summ.Column).Value = XLCellValueConverter.FromObject(value);
            }
            else if (summ.FuncNum > 0)
            {
                var funcRngAddr = calculatedRange.Column(1).RangeAddress;
                summRow.Cell(summ.Column).FormulaA1 =
                    (string.IsNullOrWhiteSpace(Cell.Formula) ? string.Empty : $"{Cell.Formula} & ")
                    + $"Subtotal({summ.FuncNum},{funcRngAddr.ToStringRelative()})";
            }
            else
            {
                throw new NotSupportedException($"Aggregate function {summ.FuncName} not supported.");
            }
        }

        private SubtotalSummaryFunc GetFunc(ProcessingContext context)
        {
            var func = new SubtotalSummaryFunc(Name, Column);
            if (HasParameter("Over"))
                func.GetCalculateDelegate = type =>
                {
                    var par = Expression.Parameter(type, "item");
                    return context.Evaluator.ParseExpression(GetParameter("Over"), new[] { par });
                    //return XLDynamicExpressionParser.ParseLambda(new[] {par}, null, GetParameter("Over"));
                };
            func.DataSource = DataSource;

            if (HasParameter("Default"))
            {
                var dlg = context.Evaluator.ParseExpression(GetParameter("Default"), new ParameterExpression[] {});
                
                func.DefaultValueForEmptySource = dlg.DynamicInvoke();
            }

            return func;
        }
    }
}
