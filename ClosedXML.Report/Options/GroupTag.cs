/*
 -----------------------------------------------------------------------
"Group"         "\Desc"
                "\Collapse"
                "\MergeLabels" *
                "\MergeLabels2" *
                "\PlaceToColumn=n" *
                "\WithHeader" *
                "\Disablesubtotals" *
                "\DisableOutline" *
                "\PageBreaks" *
-----------------------------------------------------------------------
 */

using System;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Options
{
    public class GroupTag : SortTag
    {
        public override byte Priority { get { return 200; } }

        public bool PageBreaks { get { return Parameters.ContainsKey("pagebreaks"); } }
        public bool DisableSubtotals { get { return Parameters.ContainsKey("disablesubtotals"); } }
        public bool Collapse { get { return Parameters.ContainsKey("collapse"); } }
        public bool DisableOutLine { get { return Parameters.ContainsKey("disableoutline"); } }
        public bool OutLine { get { return !Parameters.ContainsKey("disableoutline"); } }
        public int LabelToColumn { get { return Parameters.ContainsKey("placetocolumn") ? Parameters["placetocolumn"].AsInt(1) : Column; } }
        public string LabelFormat { get { return Parameters.ContainsKey("labelformat") ? Parameters["labelformat"] : null; } }
        public int Level { get; set; }

        private MergeMode? _mergeLabels;
        public MergeMode MergeLabels
        {
            get
            {
                if (_mergeLabels.HasValue)
                    return _mergeLabels.Value;

                _mergeLabels = (MergeMode)(_mergeLabels = Parameters.ContainsKey("mergelabels") ? Parameters["mergelabels"].AsEnum(MergeMode.Merge1) : MergeMode.None);
                var mergeLabels = IsWithHeader && _mergeLabels == MergeMode.None
                    ? MergeMode.Merge1
                    : (MergeMode)_mergeLabels;
                _mergeLabels = mergeLabels;
                return mergeLabels;
            }
        }

        private bool? _isWithHeader;
        public bool IsWithHeader
        {
            get { return (bool)(_isWithHeader ?? (_isWithHeader = Parameters.ContainsKey("withheader"))); }
        }

        public override void Execute(ProcessingContext context)
        {
            var fields = List.GetAll<GroupTag>()
                .OrderBy(x => context.Range.Cell(x.Cell.Row, x.Cell.Column).WorksheetColumn().ColumnNumber()).ToArray();

            var funcs = List.GetAll<SummaryFuncTag>().ToArray();
            funcs.ForEach(x => x.DataSource = (DataSource)context.Value);

            var isWithHeader = fields.Any(x => x.IsWithHeader);
            bool summaryAbove = !isWithHeader && List.HasTag("summaryabove");

            // sort range
            base.Execute(context);

            Process(context.Range, fields, summaryAbove, funcs.Select(x => x.GetFunc()).ToArray(), List.HasTag("disablegrandtotal"));

            foreach (var tag in fields.Cast<OptionTag>().Union(funcs))
            {
                tag.Enabled = false;
            }
        }

        private void Process(IXLRange root, GroupTag[] groups, bool summaryAbove, SubtotalSummaryFunc[] funcs, bool disableGrandTotal)
        {
            var groupRow = root.LastRow();
            //   DoGroups

            var level = 0;
            var r = root.Offset(0, 0, root.RowCount() - 1, root.ColumnCount());

            using (var subtotal = new Subtotal(r, summaryAbove))
            {
                if (!disableGrandTotal)
                {
                    var total = subtotal.AddGrandTotal(funcs);
                    total.SummaryRow.Cell(2).Value = total.SummaryRow.Cell(1).Value;
                    total.SummaryRow.Cell(1).Value = null;
                    level++;
                }

                foreach (var g in groups.OrderBy(x=>x.Column))
                {
                    Func<string, string> labFormat = null;
                    if (!string.IsNullOrEmpty(g.LabelFormat))
                        labFormat = title => string.Format(LabelFormat, title);

                    subtotal.GroupBy(g.Column, g.DisableSubtotals ? new SubtotalSummaryFunc[0] : funcs, g.PageBreaks, labFormat);

                    g.Level = ++level;
                }
                foreach (var g in groups.Where(x=>x.IsWithHeader).OrderBy(x=>x.Column))
                    subtotal.AddHeaders(g.Column);

                var gr = groups.Union(new[] { new GroupTag { Column = 1, Level = 1 } })
                    .ToDictionary(x => x.Level, x => x);

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
            foreach (var cell in groupRow.Cells(c => c.HasFormula))
            {
                subGroup.SummaryRow.Cell(cell.Address.ColumnNumber - groupRow.RangeAddress.FirstAddress.ColumnNumber + 1).Value = cell;
            }
            subGroup.SummaryRow.Clear(XLClearOptions.AllFormats);
            subGroup.SummaryRow.CopyStylesFrom(groupRow);
            subGroup.SummaryRow.CopyConditionalFormatsFrom(groupRow);
        }

        protected virtual void GroupRender(SubtotalGroup subGroup, GroupTag grData)
        {
            var sheet = subGroup.Range.Worksheet;
            sheet.SuspendEvents();
            if (!sheet.Row(subGroup.Range.FirstRow().RowNumber()).IsHidden && grData.Collapse)
            {
                sheet.CollapseRows(grData.Level);
            }

            if (grData.DisableOutLine)
            {
                using (var rows = sheet.Rows(subGroup.Range.RangeAddress.FirstAddress.RowNumber, subGroup.Range.RangeAddress.LastAddress.RowNumber))
                {
                    rows.Ungroup();
                }
            }
            if (subGroup.Column <= 0)
                return;

            if (grData.LabelToColumn != subGroup.Column)
                subGroup.SummaryRow.Cell(grData.LabelToColumn).Value = subGroup.SummaryRow.Cell(subGroup.Column).Value;

            if (grData.MergeLabels > 0 && subGroup.Range.RowCount()>1)
            {
                using (var rng = subGroup.Range.Column(subGroup.Column))
                {
                    rng.Cell(1).Clear(XLClearOptions.AllFormats);
                    rng.Cell(1).AsRange().Unsubscribed().CopyStylesFrom(rng.FirstCellUsed(false).AsRange().Unsubscribed());
                    rng.Cell(1).AsRange().Unsubscribed().CopyConditionalFormatsFrom(rng.FirstCellUsed(false).AsRange().Unsubscribed());
                    rng.Merge();
                    rng.Value = subGroup.GroupTitle;
                }
            }
            sheet.ResumeEvents();
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