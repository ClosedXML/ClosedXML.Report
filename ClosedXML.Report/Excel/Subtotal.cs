using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using MoreLinq;

namespace ClosedXML.Report.Excel
{
    public class Subtotal : IDisposable
    {
        private IXLRange _range;
        private readonly bool _summaryAbove;
        private bool _pageBreaks;
        private Func<string, string> _getGroupLabel;
        private IXLWorksheet Sheet => _range.Worksheet;
        private IXLWorksheet _tempSheet;
        private readonly List<SubtotalGroup> _groups = new List<SubtotalGroup>();
        public string TotalLabel { get; set; } = "Total";
        public string GrandLabel { get; set; } = "Grand";

        public Subtotal(IXLRange range) : this(range, false)
        {
        }

        public Subtotal(IXLRange range, bool summaryAbove)
        {
            _range = range;
            _summaryAbove = summaryAbove;
            Sheet.Outline.SummaryVLocation = _summaryAbove ? XLOutlineSummaryVLocation.Top : XLOutlineSummaryVLocation.Bottom;
            var workbook = Sheet.Workbook;
            const string tempsheet = "__tempsheet";
            if (!workbook.Worksheets.TryGetWorksheet(tempsheet, out _tempSheet))
            {
                _tempSheet = workbook.AddWorksheet(tempsheet);
                _tempSheet.Style = Sheet.Style;
                _tempSheet.Hide();
            }
        }

        public SubtotalGroup[] Groups
        {
            get { return _groups.ToArray(); }
        }

        public SubtotalGroup AddGrandTotal(SubtotalSummaryFunc[] summaries)
        {
            if (Sheet.Row(_range.Row(2).RowNumber()).OutlineLevel == 0)
            {
                SubtotalGroup gr;
                if (_summaryAbove)
                {
                    var rangeAddress = _range.ShiftRows(1).RangeAddress;
                    _range.InsertRowsAbove(1, true);
                    gr = CreateGroup(Sheet.Range(rangeAddress), 1, 1, GrandLabel, summaries, false);
                }
                else
                {
                    var rangeAddress = _range.RangeAddress;
                    _range.InsertRowsBelow(1, true);
                    gr = CreateGroup(Sheet.Range(rangeAddress), 1, 1, GrandLabel, summaries, false);
                }
                gr.Column = 0;
                _groups.Add(gr);
                return gr;
            }
            else return null;
        }

        public void GroupBy(int groupBy, SubtotalSummaryFunc[] summaries, bool pageBreaks = false, Func<string, string> getGroupLabel = null)
        {
            _pageBreaks = pageBreaks;
            _getGroupLabel = getGroupLabel;

            var rows = Sheet.Rows(_range.RangeAddress.FirstAddress.RowNumber, _range.RangeAddress.LastAddress.RowNumber);
            var level = Math.Min(8, rows.Max(r => r.OutlineLevel) + 1);

            var grRanges = ScanRange(groupBy);
            int grCnt = grRanges.Count(x => x.Type == RangeType.DataRange);
            _range.InsertRowsBelow(grCnt, true);
            CalculateAddresses(grRanges);

            RecalculateGroups(grRanges, true);

            ArrangeRanges(grRanges);

            foreach (var moveData in grRanges)
            {
                if (moveData.Type == RangeType.DataRange)
                    _groups.Add(CreateGroup(Sheet.Range(moveData.TargetAddress), groupBy, level, moveData.GroupTitle, summaries, _pageBreaks));
            }
            ArrangePageBreaks(Groups, grRanges);
        }

        public void AddHeaders(int column)
        {
            var grRanges = _groups
                .Where(x => x.Column <= column)
                .SelectMany(x =>
                {
                    var moveDataEntries = new List<MoveData>
                    {
                        new MoveData(x.Range.RangeAddress, RangeType.DataRange, x.GroupTitle, x.Level)
                        {
                            GroupColumn = x.Column
                        }
                    };

                    if (x.SummaryRow != null)
                        moveDataEntries.Add(new MoveData(x.SummaryRow.RangeAddress, RangeType.SummaryRow, "", x.Level - 1));
                    return moveDataEntries;
                })
                .Where(x => x.Type == RangeType.SummaryRow || x.GroupColumn >= column)
                .Union(_groups.Where(x => x.HeaderRow != null).Select(x => new MoveData(x.HeaderRow.RangeAddress, RangeType.HeaderRow, "", x.Level - 1)))
                .OrderBy(x => x.SourceAddress.FirstAddress.RowNumber)
                .ToArray();

            int grCnt = grRanges.Count(x => x.Type == RangeType.DataRange);
            _range.InsertRowsBelow(grCnt, true);
            CalculateHeaders(grRanges, column);

            ArrangeRanges(grRanges);
            RecalculateGroups(grRanges, false);

            _groups
                .Where(x => x.Column == column)
                .ForEach(g =>
                {
                    g.HeaderRow = _range.Row(g.Range.RangeAddress.FirstAddress.RowNumber - _range.RangeAddress.FirstAddress.RowNumber);
                    g.HeaderRow.Clear(XLClearOptions.Contents | XLClearOptions.DataType); // ClosedXML issue 844
                    g.HeaderRow.Cell(column).Value = g.GroupTitle;
                });

            ArrangePageBreaks(Groups, new MoveData[0]);

            SetOutlineLevels(
                _groups.SelectMany(x =>
                {
                    var moveDataEntries = new List<MoveData>
                    {
                        new MoveData(x.Range.RangeAddress, RangeType.DataRange, x.GroupTitle, x.Level)
                        {
                            GroupColumn = x.Column
                        }
                    };
                    if (x.SummaryRow != null)
                        moveDataEntries.Add(new MoveData(x.SummaryRow.RangeAddress, RangeType.SummaryRow, "", x.Level - 1));
                    return moveDataEntries;
                })
                .Union(_groups.Where(x => x.HeaderRow != null).Select(x => new MoveData(x.HeaderRow.RangeAddress, RangeType.HeaderRow, "", x.Level - 1)))
                .ToArray()
            );
        }

        public void Unsubtotal()
        {
            var rows = Sheet.Rows(_range.FirstRow().RowNumber(), _range.LastRow().RowNumber());
            rows.Ungroup(true);

            IXLRangeRow row = _range.FirstRow();
            while (!row.IsEmpty())
            {
                if (row.IsSummary())
                {
                    var rowNumber = row.RowNumber();
                    row = row.RowAbove();
                    Sheet.Row(rowNumber).Delete();
                }
                row = row.RowBelow();
            }
        }

        private void SetOutlineLevels(MoveData[] grRanges)
        {
            var rows = Sheet.Rows(_range.RangeAddress.FirstAddress.RowNumber, _range.RangeAddress.LastAddress.RowNumber);
            rows.Ungroup(true);

            foreach (var moveData in grRanges)
            {
                rows = Sheet.Rows(moveData.SourceAddress.FirstAddress.RowNumber, moveData.SourceAddress.LastAddress.RowNumber);
                foreach (var row in rows)
                {
                    row.OutlineLevel = moveData.Level;
                }
            }
        }

        private void ArrangeRanges(MoveData[] grRanges)
        {
            ExpandSummariesRanges(grRanges);

            var rows = Sheet.Rows(_range.RangeAddress.FirstAddress.RowNumber, _range.RangeAddress.LastAddress.RowNumber);
            rows.Ungroup(true);

            for (int i = grRanges.Length - 1; i >= 0; i--)
            {
                var moveData = grRanges[i];

                if (moveData.Type == RangeType.DataRange)
                    MoveRange(moveData);
                else
                    MoveSummary(moveData);
            }
        }

        private void ArrangePageBreaks(SubtotalGroup[] groups, MoveData[] scannedGroups)
        {
            var firstRow = _range.RangeAddress.FirstAddress.RowNumber;
            var lastRow = _range.RangeAddress.LastAddress.RowNumber;
            var pageBreak = Sheet.PageSetup.RowBreaks;
            pageBreak.RemoveAll(x => firstRow <= x && x <= lastRow);

            pageBreak.AddRange(
                groups.Where(x => x.PageBreaks)
                    .Select(x => _summaryAbove || x.SummaryRow == null
                        ? x.Range.RangeAddress.LastAddress.RowNumber
                        : x.SummaryRow.RangeAddress.LastAddress.RowNumber)
                    .Union(
                        scannedGroups.Where(x => x.PageBreak)
                            .Select(x => x.TargetAddress.FirstAddress.RowNumber - (_summaryAbove ? 1 : 0))
                    )
                    .Distinct());
        }

        private void MoveSummary(MoveData moveData)
        {
            var srcRng = Sheet.Range(moveData.SourceAddress);
            var trgtRng = Sheet.Range(moveData.TargetAddress);
            Sheet.Row(trgtRng.RangeAddress.FirstAddress.RowNumber).OutlineLevel = moveData.Level;

            var fcell = trgtRng.FirstCell();
            if (Equals(fcell.Address, srcRng.RangeAddress.FirstAddress))
                return;

            trgtRng.Clear();
            fcell.Value = srcRng;
            srcRng.Clear(XLClearOptions.AllContents);

            foreach (var cell in trgtRng.CellsUsed(c => c.HasFormula))
            {
                cell.FormulaA1 = ShiftFormula(cell.FormulaA1, moveData.SourceAddress.FirstAddress.RowNumber - moveData.TargetAddress.FirstAddress.RowNumber);
            }
        }

        private void MoveRange(MoveData moveData)
        {
            var srcRng = Sheet.Range(moveData.SourceAddress);
            _tempSheet.Clear();
            _tempSheet.Cell(1, 1).Value = srcRng;
            srcRng.Clear(XLClearOptions.AllContents);
            Sheet.Range(moveData.TargetAddress).Clear();
            Sheet.Cell(moveData.TargetAddress.FirstAddress).Value = _tempSheet.Range(1, 1, srcRng.RowCount(), srcRng.ColumnCount());
        }

        private SubtotalGroup CreateGroup(IXLRange groupRng, int groupClmn, int level, string title, SubtotalSummaryFunc[] summaries, bool pageBreaks)
        {
            var firstRow = groupRng.RangeAddress.FirstAddress.RowNumber;
            var lastRow = groupRng.RangeAddress.LastAddress.RowNumber;
            IXLRangeRow summRow;

            if (_summaryAbove)
            {
                var fr = _range.Row(firstRow - _range.RangeAddress.FirstAddress.RowNumber + 1);
                summRow = fr.RowAbove();
                summRow.CopyStylesFrom(fr);
            }
            else
            {
                var fr = _range.Row(lastRow - _range.RangeAddress.FirstAddress.RowNumber + 1);
                summRow = fr.RowBelow();
                summRow.CopyStylesFrom(fr);
            }

            summRow.Clear(XLClearOptions.Contents | XLClearOptions.DataType); //TODO Check if the issue persists (ClosedXML issue 844)
            summRow.Cell(groupClmn).Value = _getGroupLabel != null ? _getGroupLabel(title) : title + " "+ TotalLabel;
            Sheet.Row(summRow.RowNumber()).OutlineLevel = level - 1;

            foreach (var summ in summaries)
            {
                /*if (summ.FuncNum == 0)
                {
                    summRow.Cell(summ.Column).Value = summ.Calculate(groupRng);
                }
                else */if (summ.FuncNum > 0)
                {
                    var funcRngAddr = groupRng.Column(summ.Column).RangeAddress;
                    summRow.Cell(summ.Column).FormulaA1 = $"Subtotal({summ.FuncNum},{funcRngAddr.ToStringRelative()})";
                }
                else
                {
                    throw new NotSupportedException($"Aggregate function {summ.FuncName} not supported.");
                }
            }

            var rows = Sheet.Rows(firstRow, lastRow);
            rows.ForEach(r => r.OutlineLevel = level);

            return new SubtotalGroup(level, groupClmn, title, groupRng, summRow, pageBreaks);
        }

        public SubtotalGroup[] ScanForGroups(int groupBy)
        {
            var grRanges = ScanRange(groupBy);
            var result = new List<SubtotalGroup>(grRanges.Length);
            var rows = Sheet.Rows(_range.RangeAddress.FirstAddress.RowNumber, _range.RangeAddress.LastAddress.RowNumber);
            var level = Math.Min(8, rows.Max(r => r.OutlineLevel) + 1);

            foreach (var moveData in grRanges)
            {
                if (moveData.Type == RangeType.DataRange)
                {
                    var groupRng = Sheet.Range(moveData.SourceAddress);
                    var gr = new SubtotalGroup(level, groupBy, moveData.GroupTitle, groupRng, null, false);
                    result.Add(gr);
                }
            }

            _groups.AddRange(result);
            return result.ToArray();
        }

        private MoveData[] ScanRange(int groupBy)
        {
            IXLRangeRow lastRow = null;
            string prevVal = null;
            int groupStart = 0;
            List<MoveData> groups = new List<MoveData>();

            var rows = _range.Rows();
            foreach (var row in rows)
            {
                lastRow = row;

                var val = row.Cell(groupBy).GetString();
                var isSummaryRow = row.IsSummary();

                    if (string.IsNullOrEmpty(val) && !isSummaryRow)
                    {
                        if (groupStart > 0)
                        {
                            groups.Add(CreateMoveTask(groupBy, prevVal, _range.Cell(groupStart, 1), row.RowAbove().LastCell(), RangeType.DataRange));
                        }
                        groups.Add(CreateMoveTask(groupBy, "", row.FirstCell(), row.LastCell(), RangeType.HeaderRow));
                        prevVal = null;
                        groupStart = 0;
                        continue;
                    }

                if (val != prevVal)
                {
                    if (groupStart > 0)
                    {
                        groups.Add(CreateMoveTask(groupBy, prevVal, _range.Cell(groupStart, 1), row.RowAbove().LastCell(), RangeType.DataRange));
                    }
                    prevVal = val;
                    groupStart = !isSummaryRow ? row.RangeAddress.Relative(_range.RangeAddress).FirstAddress.RowNumber : 0;
                }
                if (isSummaryRow)
                {
                    var moveData = CreateMoveTask(groupBy, "", row.FirstCell(), row.LastCell(), RangeType.SummaryRow);
                    moveData.PageBreak = Sheet.PageSetup.RowBreaks.Any(x => row.RowNumber() - (_summaryAbove ? 1 : 0) == x);
                    groups.Add(moveData);
                }
            }
            if (lastRow != null && groupStart > 0)
            {
                groups.Add(CreateMoveTask(groupBy, prevVal, _range.Cell(groupStart, 1), lastRow.LastCell(), RangeType.DataRange));
            }

            return groups.ToArray();
        }

        private MoveData CreateMoveTask(int groupColumn, string title, IXLCell firstCell, IXLCell lastCell, RangeType rangeType)
        {
            var groupRng = _range.Range(firstCell, lastCell);
            var level = firstCell.WorksheetRow().OutlineLevel;
            var group = new MoveData(groupRng.RangeAddress, rangeType, title, level) {GroupColumn = groupColumn};
            return group;
        }

        private void CalculateAddresses(MoveData[] groups)
        {
            if (!groups.Any())
                return;

            var firstRow = _range.RangeAddress.FirstAddress.RowNumber;
            var firstCol = _range.RangeAddress.FirstAddress.ColumnNumber;
            var lastCol = _range.RangeAddress.LastAddress.ColumnNumber;

            var rIdx = 0;
            foreach (var gr in groups)
            {
                if (gr.Type == RangeType.DataRange && _summaryAbove)
                    rIdx++;
                var grRowCnt = gr.SourceAddress.LastAddress.RowNumber - gr.SourceAddress.FirstAddress.RowNumber + 1;
                var trgtRng = Sheet.Range(firstRow + rIdx, firstCol, firstRow + grRowCnt + rIdx - 1, lastCol);
                gr.TargetAddress = trgtRng.RangeAddress;

                rIdx += grRowCnt;
                if (gr.Type == RangeType.DataRange && !_summaryAbove)
                    rIdx++;
            }
        }

        private void CalculateHeaders(MoveData[] groups, int column)
        {
            var firstRow = _range.RangeAddress.FirstAddress.RowNumber;
            var firstCol = _range.RangeAddress.FirstAddress.ColumnNumber;
            var lastCol = _range.RangeAddress.LastAddress.ColumnNumber;

            var rIdx = 0;
            foreach (var gr in groups)
            {
                if (gr.GroupColumn == column)
                    rIdx++;
                var grRowCnt = gr.SourceAddress.LastAddress.RowNumber - gr.SourceAddress.FirstAddress.RowNumber + 1;
                var trgtRng = Sheet.Range(firstRow + rIdx, firstCol, firstRow + grRowCnt + rIdx - 1, lastCol);
                gr.TargetAddress = trgtRng.RangeAddress;

                rIdx += grRowCnt;
            }
        }

        private void ExpandSummariesRanges(MoveData[] groups)
        {
            var sheet = _range.Worksheet;
            foreach (var gr in groups)
            {
                if (gr.Type == RangeType.SummaryRow)
                {
                    // expand summary formulas
                    var row = sheet.Row(gr.SourceAddress.FirstAddress.RowNumber);
                    foreach (var cell in row.CellsUsed(c => c.HasFormula))
                    {
                        cell.FormulaA1 = ExpandFormula(groups, cell.FormulaA1);
                    }
                }
            }
        }

        private void RecalculateGroups(MoveData[] extendedGroups, bool extendAtEnd)
        {
            foreach (var gr in extendedGroups.Where(g => g.Type == RangeType.DataRange).Reverse())
            {
                //TODO What are the proper names of these collections?
                var collectionA = _groups
                    .Where(g => g.Column > gr.GroupColumn && gr.SourceAddress.Contains(g.Range.RangeAddress))
                    .ToList();
                var collectionB = _groups
                    .Where(g => g.Column == gr.GroupColumn && Equals(g.Range.RangeAddress, gr.SourceAddress))
                    .ToList();
                var collectionC = _groups
                    .Where(g => g.Range.RangeAddress.FirstAddress.RowNumber > gr.SourceAddress.LastAddress.RowNumber)
                    .ToList();
                var collectionD = _groups
                    .Where(g => g.Column < gr.GroupColumn && g.Range.RangeAddress.Contains(gr.SourceAddress))
                    .ToList();

                if (!extendAtEnd)
                    collectionA
                        .ForEach(g =>
                        {
                            g.HeaderRow = g.HeaderRow?.ShiftRows(1)?.Row(1);
                            g.Range = g.Range.ShiftRows(1);
                            g.SummaryRow = g.SummaryRow?.ShiftRows(1)?.Row(1);
                        });

                collectionB
                    .ForEach(g =>
                    {
                        g.Range = g.Range.ShiftRows(1);
                        g.SummaryRow = g.SummaryRow?.ShiftRows(1)?.Row(1);
                    });

                collectionC
                    .ForEach(g =>
                    {
                        g.HeaderRow = g.HeaderRow?.ShiftRows(1)?.Row(1);
                        g.Range = g.Range.ShiftRows(1);
                        g.SummaryRow = g.SummaryRow?.ShiftRows(1)?.Row(1);
                    });

                collectionD
                    .ForEach(g =>
                    {
                        g.Range = g.Range.ExtendRows(1);
                        if (!_summaryAbove) g.SummaryRow = g.SummaryRow?.ShiftRows(1)?.Row(1);
                    });
            }
        }

        private string ExpandFormula(MoveData[] groups, string formula)
        {
            var pars = _range.Worksheet.GetRangeParameters(formula).Where(r => r.Key.Contains(":"));
            foreach (var addr in pars)
            {
                var firstGroup = groups.FirstOrDefault(x => x.SourceAddress.FirstAddress.RowNumber == addr.Value.FirstAddress.RowNumber);
                var lastGroup = groups.FirstOrDefault(x => x.SourceAddress.LastAddress.RowNumber == addr.Value.LastAddress.RowNumber)
                    ?? groups.FirstOrDefault(x => x.SourceAddress.LastAddress.RowNumber - 1 == addr.Value.LastAddress.RowNumber);
                if (firstGroup == null || lastGroup == null)
                    continue;

                var sheet = addr.Value.Worksheet;
                var rangeAddress = sheet.Range(
                    firstGroup.TargetAddress.FirstAddress.RowNumber,
                    addr.Value.FirstAddress.ColumnNumber,
                    lastGroup.TargetAddress.LastAddress.RowNumber,
                    addr.Value.LastAddress.ColumnNumber).RangeAddress;
                formula = formula.Replace(addr.Key, rangeAddress.ToStringRelative());

            }
            return formula;
        }

        private string ShiftFormula(string formula, int rowCount)
        {
            var pars = _range.Worksheet.GetRangeParameters(formula).Where(r => r.Key.Contains(":"));
            foreach (var addr in pars)
            {
                var sheet = addr.Value.Worksheet;
                var rangeAddress = sheet.Range(
                    addr.Value.FirstAddress.RowNumber + rowCount,
                    addr.Value.FirstAddress.ColumnNumber,
                    addr.Value.LastAddress.RowNumber + rowCount,
                    addr.Value.LastAddress.ColumnNumber).RangeAddress;
                formula = formula.Replace(addr.Key, rangeAddress.ToStringRelative());
            }
            return formula;
        }

        private class MoveData
        {
            public MoveData(IXLRangeAddress sourceAddress, RangeType type, string groupTitle, int level)
            {
                SourceAddress = sourceAddress;
                Type = type;
                GroupTitle = groupTitle;
                Level = level;
            }

            public IXLRangeAddress SourceAddress { get; }
            public IXLRangeAddress TargetAddress { get; set; }
            public RangeType Type { get; }
            public string GroupTitle { get; }
            public int GroupColumn { get; set; }
            public int Level { get; }
            public bool PageBreak { get; set; }
        }

        public enum RangeType
        {
            HeaderRow,
            DataRange,
            SummaryRow
        }

        public void Dispose()
        {
            if (_tempSheet != null)
            {
                _tempSheet.Delete();
                _tempSheet = null;
            }
        }
    }
}
