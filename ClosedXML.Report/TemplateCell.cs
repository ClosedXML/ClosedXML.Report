using ClosedXML.Excel;

namespace ClosedXML.Report
{
    public class TemplateCell
    {
        public TemplateCellType CellType { get; internal set; }
        public bool IsCalculated { get; }
        public int Row { get; internal set; }
        public int Column { get; internal set; }
        public string Formula { get; set; }
        public object Value { get; set; }
        public IXLCell XLCell { get; internal set; }

        public TemplateCell()
        {
        }

        public TemplateCell(int row, int column, IXLCell xlCell)
            : this()
        {
            XLCell = xlCell;
            Row = row;
            Column = column;
            if (xlCell != null)
            {
                Value = xlCell.Value;
                if (xlCell.HasFormula)
                {
                    Formula = xlCell.FormulaA1;
                    CellType = TemplateCellType.Formula;
                }
                else
                {
                    CellType = TemplateCellType.Value;
                    var strVal = xlCell.GetString();
                        if (strVal.StartsWith("&="))
                        {
                            CellType = TemplateCellType.Formula;
                            Formula = strVal.Substring(2);
                        }
                    if (strVal.Contains("{{"))
                    {
                        IsCalculated = true;
                    }
                }
            }
        }

        public string GetString()
        {
            switch (CellType)
            {
                case TemplateCellType.Value: return Value.ToString();
                case TemplateCellType.Formula: return Formula;
                default: return "";
            }
        }

        public IXLCell GetXlCell(IXLRange range)
        {
            return range.Cell(Row, Column);
        }

        public TemplateCell Clone()
        {
            return (TemplateCell) MemberwiseClone();
        }
    }

    public enum TemplateCellType
    {
        None,
        NewRow,
        Formula,
        Value
    }
}
