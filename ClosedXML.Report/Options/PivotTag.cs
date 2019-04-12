/*
PivotTable Options Package
================================================
OPTION          PARAMS                OBJECTS   
================================================
"Pivot"          "Name="              Range     
                 "Dst="
                 "RowGrand"
                 "ColumnGrand"
                 "NoPreserveFormatting"
                 "CaptionNoFormatting"
                 "MergeLabels"
                 "ShowButtons"
                 "TreeLayout"
                 "AutofitColumns"
                 "NoSort"

"Data"                                Column    
"Row"                                 Column    
"Column"                              Column    
"Page"                                Column    
================================================
*/

using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;
using MoreLinq;

namespace ClosedXML.Report.Options
{
    public class PivotTag: OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var fields = List.GetAll<PivotTag>()
                .OrderBy(x => context.Range.Cell(x.Cell.Row, x.Cell.Column).WorksheetColumn().ColumnNumber()).ToArray();

            // Init variables
            var wb = context.Range.Worksheet.Workbook;
            var rowTags = List.GetAll(new[] {"row"}).OrderBy(t => t.Column);
            var colTags = List.GetAll(new[] {"column", "col"}).OrderBy(t => t.Column);
            var pageTags = List.GetAll(new[] {"page"}).OrderBy(t => t.Column).ToList();
            var dataTags = List.GetAll<DataPivotTag>().OrderBy(t => t.Column);

            var pivotTag = fields.FirstOrDefault(t => t.Name.ToLower() == "pivot") ?? this;
            var (tableName, dstSheet, dstCell) = GetDestination(pivotTag, wb, pageTags);
            var pt = CreatePivot(pivotTag, context, dstSheet, tableName, dstCell);

            foreach (var optionTag in pageTags)
            {
                var tag = (FieldPivotTag) optionTag;
                var colName = GetColumnName(context, tag);
                var field = pt.ReportFilters.Add(colName);
                field.ShowBlankItems = false;
                BuildFormatting(pivotTag, tag, field);
            }

            foreach (var optionTag in rowTags)
            {
                var tag = (FieldPivotTag) optionTag;
                var colName = GetColumnName(context, tag);
                var field = pt.RowLabels.Add(colName);
                field.ShowBlankItems = false;
                BuildFormatting(pivotTag, tag, field);
            }

            foreach (var optionTag in colTags)
            {
                var tag = (FieldPivotTag) optionTag;
                var colName = GetColumnName(context, tag);
                var field = pt.ColumnLabels.Add(colName);
                field.ShowBlankItems = false;
                BuildFormatting(pivotTag, tag, field);
            }

            // Build data fields (datarange)
            foreach (var tag in dataTags)
            {
                var colName = GetColumnName(context, tag);
                var field = pt.Values.Add(colName);
                field.SummaryFormula = tag.SummaryFormula;
                BuildFormatting(pivotTag, tag, field);
            }

            var tags = fields.Union<OptionTag>(List.GetAll<SummaryFuncTag>());
            tags.ForEach(tag => tag.Enabled = false);
        }

        private void BuildFormatting(PivotTag pivotTag, FieldPivotTag tag, IXLPivotField pf)
        {
            foreach (var func in tag.SubtotalFunction)
            {
                pf.AddSubtotal(func);
            }

            // TODO pivot field formatting
            /*
               '    Rem Build page fields' + vbCR +
              '    For i = 1 To PageFieldsCount' + vbCR +
              '      V = Pages(i)' + vbCR +
              '      Set PF = PT.PivotFields(V(2))' + vbCR +
              '      PF.Subtotals = V(3)' + vbCR +
              '      If Args(14) = True Then' + vbCR +
              '          Set FmtRange = SrcRange.Cells(2, V(1) - 1)' + vbCR +
              '          If (PF.DataType = xlDate) Or (PF.DataType = xlNumber) Then' + vbCR +
              '            PF.NumberFormat = FmtRange.NumberFormat' + vbCR +
              '          End If' + vbCR +
              '          PF.DataRange.Interior.ColorIndex = FmtRange.Interior.ColorIndex' + vbCR +
              '          PF.DataRange.Font.Name = FmtRange.Font.Name' + vbCR +
              '          PF.DataRange.Font.Color = FmtRange.Font.Color' + vbCR +
              '          PF.DataRange.Font.Size = FmtRange.Font.Size' + vbCR +
              '          PF.DataRange.Font.FontStyle = FmtRange.Font.FontStyle' + vbCR +
              '          PF.DataRange.HorizontalAlignment = FmtRange.HorizontalAlignment' + vbCR +
              '          PF.DataRange.VerticalAlignment = FmtRange.VerticalAlignment' + vbCR +
              '      End If' + vbCR +
              '    Next' + vbCR +
            */
        }

        private void BuildFormatting(PivotTag pivot, DataPivotTag tag, IXLPivotValue pf)
        {
            if (pivot.HasParameter("NoPreserveFormatting") && pivot.HasParameter("CaptionNoFormatting"))
            {
                var fmtRange = tag.Cell;
            }
            // TODO pivot value formatting
            /*
              '    If (Args(14) = True) And (!HasParameter("CaptionNoFormatting")) Then' + vbCR +
              '      On Error Resume Next' + vbCR +
              '      For i = 1 To DataFieldsCount' + vbCR +
              '        V = Datas(i)' + vbCR +
              '        Set PF = PT.DataFields(i)' + vbCR + // V(2) & " ")' + vbCR +
              '        Set FmtRange = SrcRange.Cells(2, V(1) - 1)' + vbCR +
              '        If PF.DataType <> xlText Then' + vbCR +
              '          PF.NumberFormat = FmtRange.NumberFormat' + vbCR +
              '        End If' + vbCR +
              '        PF.DataRange.Interior.ColorIndex = FmtRange.Interior.ColorIndex' + vbCR +
              '        PF.DataRange.Font.Name = FmtRange.Font.Name' + vbCR +
              '        PF.DataRange.Font.Color = FmtRange.Font.Color' + vbCR +
              '        PF.DataRange.Font.Size = FmtRange.Font.Size' + vbCR +
              '        PF.DataRange.Font.FontStyle = FmtRange.Font.FontStyle' + vbCR +
              '        PF.DataRange.HorizontalAlignment = FmtRange.HorizontalAlignment' + vbCR +
              '        PF.DataRange.VerticalAlignment = FmtRange.VerticalAlignment' + vbCR +
              '        if (!HasParameter("CaptionNoFormatting"))' + vbCR +
              '          Set CaptionRange = SrcRange.Cells(1, V(1) - 1)' + vbCR +
              '          PF.LabelRange.Interior.ColorIndex = CaptionRange.Interior.ColorIndex' + vbCR +
              '          PF.LabelRange.Font.Name = CaptionRange.Font.Name' + vbCR +
              '          PF.LabelRange.Font.Color = CaptionRange.Font.Color' + vbCR +
              '          PF.LabelRange.Font.Size = CaptionRange.Font.Size' + vbCR +
              '          PF.LabelRange.Font.FontStyle = CaptionRange.Font.FontStyle' + vbCR +
              '          PF.LabelRange.HorizontalAlignment = CaptionRange.HorizontalAlignment' + vbCR +
              '          PF.LabelRange.VerticalAlignment = CaptionRange.VerticalAlignment' + vbCR +
              '        End If' + vbCR +
              '      Next' + vbCR +
              '    End If' + vbCR +
            */
        }

        private (string, IXLWorksheet, IXLCell) GetDestination(PivotTag pivot, XLWorkbook wb, IEnumerable<OptionTag> pageTags)
        {
            string tableName = pivot.GetParameter("name");
            if (tableName.IsNullOrWhiteSpace()) tableName = "PivotTable";

            IXLWorksheet dstSheet;
            IXLCell dstCell;
            var dstStr = pivot.GetParameter("dst");
            if (!dstStr.IsNullOrWhiteSpace())
            {
                var shSp = dstStr.IndexOf("!", StringComparison.Ordinal);
                dstSheet = wb.Worksheet(dstStr.Substring(0, shSp));
                if (dstSheet == null)
                    throw new TemplateParseException($"Can\'t find pivot destination sheet \'{dstStr.Substring(0, shSp)}\'", Cell.XLCell.AsRange());
                dstStr = dstStr.Substring(shSp+1, dstStr.Length - shSp - 1);
                dstCell = dstSheet.Cell(dstStr);
                if (dstCell == null)
                    throw new TemplateParseException($"Can\'t find pivot destination cell \'{dstStr}\'", Cell.XLCell.AsRange());
            }
            else
            {
                dstSheet = wb.AddWorksheet(tableName);
                dstCell = dstSheet.Cell(pageTags.Count() + 3, 2);
            }
            return (tableName, dstSheet, dstCell);
        }

        private IXLPivotTable CreatePivot(PivotTag pivot, ProcessingContext context, IXLWorksheet targetSheet, string tableName, IXLCell targetCell)
        {
            var rowOffset = context.Range.RangeAddress.FirstAddress.RowNumber > 1 ? -1 : 0;
            IXLRange srcRange = context.Range.Offset(rowOffset, 1, context.Range.RowCount(), context.Range.ColumnCount() - 1);
            var pt = targetSheet.PivotTables.Add(tableName, targetCell, srcRange);
            pt.MergeAndCenterWithLabels = pivot.HasParameter("MergeLabels");
            pt.ShowExpandCollapseButtons = pivot.HasParameter("ShowButtons");
            pt.ClassicPivotTableLayout = !pivot.HasParameter("TreeLayout");
            pt.AutofitColumns = pivot.HasParameter("AutofitColumns");
            pt.SortFieldsAtoZ = !pivot.HasParameter("NoSort");
            pt.PreserveCellFormatting = !pivot.HasParameter("NoPreserveFormatting");
            pt.ShowGrandTotalsColumns = pivot.HasParameter("ColumnGrand");
            pt.ShowGrandTotalsRows = pivot.HasParameter("RowGrand");
            pt.SaveSourceData = true;
            pt.FilterAreaOrder = XLFilterAreaOrder.DownThenOver;
            pt.RefreshDataOnOpen = true;
            return pt;
        }

        private static string GetColumnName(ProcessingContext context, OptionTag tag)
        {
            return context.Range.FirstRow().RowAbove().Cell(tag.Column).GetString();
        }
    }

    public class FieldPivotTag: PivotTag
    {
        public IEnumerable<XLSubtotalFunction> SubtotalFunction
        {
            get
            {
                return List.GetAll<SummaryFuncTag>().Where(x => x.Column == Column).Select(GetSubtotalFunction);
            }
        }

        private XLSubtotalFunction GetSubtotalFunction(SummaryFuncTag tag)
        {
            var upper = tag.Name.ToUpper();
            if (upper == "SUM")
                return XLSubtotalFunction.Sum;
            else if (upper == "AVG" || upper == "AVERAGE")
                return XLSubtotalFunction.Average;
            else if (upper == "COUNT")
                return XLSubtotalFunction.Count;
            else if (upper == "COUNTNUMS")
                return XLSubtotalFunction.CountNumbers;
            else if (upper == "MAX")
                return XLSubtotalFunction.Maximum;
            else if (upper == "MIN")
                return XLSubtotalFunction.Minimum;
            else if (upper == "PRODUCT")
                return XLSubtotalFunction.Product;
            else if (upper == "STDEV")
                return XLSubtotalFunction.StandardDeviation;
            else if (upper == "STDEVP")
                return XLSubtotalFunction.PopulationStandardDeviation;
            else if (upper == "VAR")
                return XLSubtotalFunction.Variance;
            else if (upper == "VARP")
                return XLSubtotalFunction.PopulationVariance;
            else
                return XLSubtotalFunction.Count;
        }
    }

    public class DataPivotTag : PivotTag
    {
        public XLPivotSummary SummaryFormula
        {
            get
            {
                var sumtag = List.GetAll<SummaryFuncTag>().First(x => x.Column == Column);
                var upper = sumtag.Name.ToUpper();
                if (upper =="SUM")
                    return XLPivotSummary.Sum;
                else if (upper =="AVG" || upper =="AVERAGE")
                    return XLPivotSummary.Average;
                else if (upper =="COUNT")
                    return XLPivotSummary.Count;
                else if (upper =="COUNTNUMS")
                    return XLPivotSummary.CountNumbers;
                else if (upper =="MAX")
                    return XLPivotSummary.Maximum;
                else if (upper =="MIN")
                    return XLPivotSummary.Minimum;
                else if (upper =="PRODUCT")
                    return XLPivotSummary.Product;
                else if (upper =="STDEV")
                    return XLPivotSummary.StandardDeviation;
                else if (upper =="STDEVP")
                    return XLPivotSummary.PopulationStandardDeviation;
                else if (upper =="VAR")
                    return XLPivotSummary.Variance;
                else if (upper =="VARP")
                    return XLPivotSummary.PopulationVariance;
                else
                    return XLPivotSummary.Count;
            }
        }
    }
}
