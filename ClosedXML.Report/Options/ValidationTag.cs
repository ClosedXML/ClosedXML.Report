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

namespace ClosedXML.Report.Options
{
    public class ValidationTag : OptionTag
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
                    xlCell.GetDataValidation().List(listRange);
                else
                    xlCell.GetDataValidation().List(listStr);
            }

            xlCell.GetDataValidation().AllowedValues = GetParameter("AllowedValues").AsEnum(XLAllowedValues.AnyValue);
            xlCell.GetDataValidation().Operator = GetParameter("Operator").AsEnum(XLOperator.EqualTo);
            xlCell.GetDataValidation().Value = GetParameter("Value");
            xlCell.GetDataValidation().MinValue = GetParameter("MinValue");
            xlCell.GetDataValidation().MaxValue = GetParameter("MaxValue");
            xlCell.GetDataValidation().IgnoreBlanks = !GetParameter("ProcessBlanks").AsBool();
            xlCell.GetDataValidation().ShowErrorMessage = !GetParameter("HideErrorMessage").AsBool();
            xlCell.GetDataValidation().ErrorTitle = GetParameter("ErrorTitle");
            xlCell.GetDataValidation().ErrorMessage = GetParameter("ErrorMessage");
            xlCell.GetDataValidation().ShowInputMessage = !GetParameter("HideInputMessage").AsBool();
            xlCell.GetDataValidation().InputTitle = GetParameter("InputTitle");
            xlCell.GetDataValidation().InputMessage = GetParameter("InputMessage");
            xlCell.GetDataValidation().InCellDropdown = !GetParameter("HideDropdown").AsBool();*/
        }
    }
}
