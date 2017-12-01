using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace ClosedXML.Report
{
    internal class TemplateCells : List<TemplateCell>
    {
        public TemplateCells(RangeTemplate template)
        {
            Template = template;
        }

        public RangeTemplate Template { get; private set; }

        public TemplateCell Add(int row, int column, IXLCell xlCell)
        {
            var result = new TemplateCell(row, column, xlCell);
            base.Add(result);
            return result;
        }

        internal TemplateCell AddNewRow()
        {
            var result = new TemplateCell();
            base.Add(result);
            result.CellType = TemplateCellType.NewRow;
            return result;
        }

        public TemplateCell FirstCellOfRow(int row)
        {
            return this.FirstOrDefault(cell => cell.Row == row);
        }
    }
}