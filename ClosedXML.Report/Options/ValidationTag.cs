/*
Validation Options
====================================================
PARAMS                  
====================================================
AllowedValues       AnyValue(default)|WholeNumber|Decimal|Date|Time|TextLength|List|Custom
Operator            EqualTo(default)|NotEqualTo|GreaterThan|LessThan|EqualOrGreaterThan|EqualOrLessThan|Between|NotBetween
Value               String|Range
MinValue            String|Range
MaxValue            String|Range
ProcessBlanks       on|off(default)
HideErrorMessage    on|off(default)
ErrorTitle          String
ErrorMessage        String
HideInputMessage    on|off(default)
InputTitle          String
InputMessage        String
HideDropdown        on|off(default)
================================================

EXAMPLES
================================================
<<Validation AllowedValues="WholeNumber" Operator="Between" MinValue="=D11" MaxValue="100" ProcessBlanks>>
<<Validation List="=D11:D18">>
================================================
*/

using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class ValidationTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            /*var xlCell = Cell.GetXlCell(context.Range);
            var ws = Range.Worksheet;

            if (HasParameter("List"))
            {
                var listStr = GetParameter("List");
                var listRange = ws.ParseRange(listStr);
                if (listRange != null)
                    xlCell.DataValidation.List(listRange);
                else
                    xlCell.DataValidation.List(listStr);
            }

            xlCell.DataValidation.AllowedValues = GetParameter("AllowedValues").AsEnum(XLAllowedValues.AnyValue);
            xlCell.DataValidation.Operator = GetParameter("Operator").AsEnum(XLOperator.EqualTo);
            xlCell.DataValidation.Value = GetParameter("Value");
            xlCell.DataValidation.MinValue = GetParameter("MinValue");
            xlCell.DataValidation.MaxValue = GetParameter("MaxValue");
            xlCell.DataValidation.IgnoreBlanks = !GetParameter("ProcessBlanks").AsBool();
            xlCell.DataValidation.ShowErrorMessage = !GetParameter("HideErrorMessage").AsBool();
            xlCell.DataValidation.ErrorTitle = GetParameter("ErrorTitle");
            xlCell.DataValidation.ErrorMessage = GetParameter("ErrorMessage");
            xlCell.DataValidation.ShowInputMessage = !GetParameter("HideInputMessage").AsBool();
            xlCell.DataValidation.InputTitle = GetParameter("InputTitle");
            xlCell.DataValidation.InputMessage = GetParameter("InputMessage");
            xlCell.DataValidation.InCellDropdown = !GetParameter("HideDropdown").AsBool();*/
        }
    }
}
