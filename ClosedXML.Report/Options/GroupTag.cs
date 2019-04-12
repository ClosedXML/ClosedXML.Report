/*
=======================================================================
OPTION          PARAMS                OBJECTS      RNG     Priority
=======================================================================
"Group"         "\Desc"               Range        rD      Higher
                "\Collapse"
                "\MergeLabels=[Merge1|Merge2|Merge3]"
                "\PlaceToColumn=n"
                "\WithHeader"
                "\Disablesubtotals"
                "\DisableOutline"
                "\PageBreaks"
                "\TotalLabel"
                "\GrandLabel"

"SummaryAbove"                        Range        rD      Normal

"DisableGrandTotal"                   Range        rD      Normal
-----------------------------------------------------------------------
  RNG:
    r - range
    t - root range
    m - master range
    d - detail range 
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
    public class GroupTag : SortTag
    {
        private int _maxLevel;

        public bool PageBreaks => Parameters.ContainsKey("pagebreaks");
        public bool DisableSubtotals => Parameters.ContainsKey("disablesubtotals");
        public bool Collapse => Parameters.ContainsKey("collapse");
        public bool DisableOutLine => Parameters.ContainsKey("disableoutline");
        public bool OutLine => !Parameters.ContainsKey("disableoutline");
        public int LabelToColumn => Parameters.ContainsKey("placetocolumn") ? Parameters["placetocolumn"].AsInt(1)+1 /* +1 for special column*/  : Column;
        public string LabelFormat => Parameters.ContainsKey("labelformat") ? Parameters["labelformat"] : null;
        public string TotalLabel => Parameters.ContainsKey("totallabel") ? Parameters["totallabel"] : null;
        public string GrandLabel => Parameters.ContainsKey("grandlabel") ? Parameters["grandlabel"] : null;
        public int Level { get; set; }

        private MergeMode? _mergeLabels;
        public MergeMode MergeLabels
        {
            get
            {
                if (_mergeLabels.HasValue)
                    return _mergeLabels.Value;

                _mergeLabels = (MergeMode)(_mergeLabels = Parameters.ContainsKey("mergelabels") ? Parameters["mergelabels"].AsEnum(MergeMode.Merge1) : MergeMode.None);
                if (IsWithHeader && _mergeLabels == MergeMode.None)
                    _mergeLabels = MergeMode.Merge1;
                return _mergeLabels.Value;
            }
        }

        private bool? _isWithHeader;
        public bool IsWithHeader => (bool)(_isWithHeader ?? (_isWithHeader = Parameters.ContainsKey("withheader")));

        public override void Execute(ProcessingContext context)
        {
            if (!(context.Value is DataSource))
            {
                var xlCell = Cell.GetXlCell(context.Range);
                xlCell.Value = "The GROUP tag can't be used outside the named range.";
                xlCell.Style.Font.FontColor = XLColor.Red;
                throw new TemplateParseException("The GROUP tag can't be used outside the named range.", xlCell.AsRange());
            }

            var fields = List.GetAll<GroupTag>()
                .OrderBy(x => context.Range.Cell(x.Cell.Row, x.Cell.Column).WorksheetColumn().ColumnNumber()).ToArray();

            var funcs = List.GetAll<SummaryFuncTag>().ToArray();
            funcs.ForEach(x => x.DataSource = (DataSource)context.Value);
            if (!funcs.Any())
                Parameters["totallabel"] = "";

            var isWithHeader = fields.Any(x => x.IsWithHeader);
            bool summaryAbove = !isWithHeader && List.HasTag("summaryabove");

            // sort range
            base.Execute(context);

            Process(context.Range, fields, summaryAbove, funcs.Select(x => x.GetFunc()).ToArray(), DisableGrandTotal);

            foreach (var tag in fields.Cast<OptionTag>().Union(funcs))
            {
                tag.Enabled = false;
            }
        }

        private bool DisableGrandTotal => List.HasTag("disablegrandtotal") || !List.GetAll<SummaryFuncTag>().Any();

        private void Process(IXLRange root, GroupTag[] groups, bool summaryAbove, SubtotalSummaryFunc[] funcs, bool disableGrandTotal)
        {
            var groupRow = root.LastRow();
            //   DoGroups

            var level = 0;
            var r = root.Offset(0, 0, root.RowCount() - 1, root.ColumnCount());

            using (var subtotal = new Subtotal(r, summaryAbove))
            {
                if (TotalLabel != null) subtotal.TotalLabel = TotalLabel;
                if (GrandLabel != null) subtotal.GrandLabel = GrandLabel;
                if (!disableGrandTotal)
                {
                    var total = subtotal.AddGrandTotal(funcs);
                    total.SummaryRow.Cell(2).Value = total.SummaryRow.Cell(1).Value;
                    total.SummaryRow.Cell(1).Value = null;
                    level++;
                }

                foreach (var g in groups.OrderBy(x => x.Column))
                {
                    Func<string, string> labFormat = null;
                    if (!string.IsNullOrEmpty(g.LabelFormat))
                        labFormat = title => string.Format(LabelFormat, title);

                    if (g.MergeLabels == MergeMode.Merge2 && funcs.Length == 0)
                        subtotal.ScanForGroups(g.Column);
                    else
                        subtotal.GroupBy(g.Column, g.DisableSubtotals ? new SubtotalSummaryFunc[0] : funcs, g.PageBreaks, labFormat);

                    g.Level = ++level;
                }

                _maxLevel = level;
                foreach (var g in groups.Where(x=>x.IsWithHeader).OrderBy(x=>x.Column))
                    subtotal.AddHeaders(g.Column);

                Dictionary<int, GroupTag> gr;
                if (disableGrandTotal)
                {
                    gr = groups.ToDictionary(x => x.Level, x => x);
                }
                else
                {
                    gr = groups.Union(new[] { new GroupTag { Column = 1, Level = 1 } })
                        .ToDictionary(x => x.Level, x => x);
                }

                foreach (var subGroup in subtotal.Groups.OrderBy(x=>x.Column).Reverse())
                {
                    var groupTag = gr[subGroup.Level];
                    FormatHeaderFooter(subGroup, groupRow);

                    GroupRender(subGroup, groupTag);
                }
            }
            //   Rem DoDeleteSpecialRow
            root.LastRow().Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        private static void FormatHeaderFooter(SubtotalGroup subGroup, IXLRangeRow groupRow)
        {
            if (subGroup.HeaderRow != null)
            {
                subGroup.HeaderRow.Clear(XLClearOptions.AllFormats);
                subGroup.HeaderRow.CopyStylesFrom(groupRow);
                subGroup.HeaderRow.CopyConditionalFormatsFrom(groupRow);

            }

            if (subGroup.SummaryRow != null)
            {
                foreach (var cell in groupRow.Cells(c => c.HasFormula))
                {
                    subGroup.SummaryRow.Cell(cell.Address.ColumnNumber - groupRow.RangeAddress.FirstAddress.ColumnNumber + 1).Value = cell;
                }

                subGroup.SummaryRow.Clear(XLClearOptions.AllFormats);
                subGroup.SummaryRow.CopyStylesFrom(groupRow);
                subGroup.SummaryRow.CopyConditionalFormatsFrom(groupRow);
            }
        }

        protected virtual void GroupRender(SubtotalGroup subGroup, GroupTag grData)
        {
            var sheet = subGroup.Range.Worksheet;
            if (!sheet.Row(subGroup.Range.FirstRow().RowNumber()).IsHidden && grData.Collapse)
            {
                sheet.CollapseRows(grData.Level);
            }

            if (grData.DisableOutLine)
            {
                sheet.Rows(subGroup.Range.RangeAddress.FirstAddress.RowNumber, subGroup.Range.RangeAddress.LastAddress.RowNumber)
                    .Ungroup();
            }

            if (subGroup.Column <= 0)
                return;

            if (grData.LabelToColumn != subGroup.Column && subGroup.SummaryRow != null)
                subGroup.SummaryRow.Cell(grData.LabelToColumn).Value = subGroup.SummaryRow.Cell(subGroup.Column).Value;

            if (grData.MergeLabels > 0)
            {
                var rng = subGroup.Range.Column(subGroup.Column);
                if (subGroup.Range.RowCount() > 1)
                {
                    int cellIdx = _maxLevel - subGroup.Level + 1;
                    var style = rng.Cell(cellIdx).Style;
                    rng.Merge();
                    rng.Style = style;
                    rng.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    if (grData.MergeLabels != MergeMode.Merge2)
                        rng.Cell(1).Value = "";
                }
                else
                {
                    if (grData.MergeLabels != MergeMode.Merge2)
                        rng.Cell(1).Value = "";
                }
            }
        }

        public enum MergeMode
        {
            None,
            Merge1,
            Merge2,
            Merge3,
        }
    }
}
