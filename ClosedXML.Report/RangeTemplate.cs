using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report
{
    public class RangeTemplate
    {
        private RangeTemplate[] _subranges;
        private int _rowCnt;
        private readonly int _colCnt;
        private IXLRange[] _mergedRanges;
        private readonly TemplateCells _cells;
        private readonly TagsList _tags;
        private readonly TagsList _rangeTags;
        private readonly TagsEvaluator _tagsEvaluator;
        private RangeOptionTag _rangeOption;
        private TempSheetBuffer _buff;
        private IXLRange _rowRange;
        private IXLRangeRow _optionsRow;
        private bool _optionsRowIsEmpty = true;
        private IXLConditionalFormat[] _condFormats;
        private IXLConditionalFormat[] _totalsCondFormats;

        public string Source { get; private set; }
        public string Name { get; private set; }

        internal RangeTemplate(IXLNamedRange range, TempSheetBuffer buff)
        {
            _rowRange = range.Ranges.First();
            _cells = new TemplateCells(this);
            _tagsEvaluator = new TagsEvaluator();
            var wb = _rowRange.Worksheet.Workbook;
            _buff = buff;
            _tags = new TagsList();
            _rangeTags = new TagsList();
            Name = range.Name;
            Source = range.Name;
            wb.NamedRanges.Add(range.Name + "_tpl", range.Ranges);
        }

        internal RangeTemplate(IXLNamedRange range, TempSheetBuffer buff, int rowCnt, int colCnt) : this(range, buff)
        {
            _rowCnt = rowCnt;
            _colCnt = colCnt;
        }


        public static RangeTemplate Parse(IXLNamedRange range)
        {
            var wb = range.Ranges.First().Worksheet.Workbook;
            return Parse(range, new TempSheetBuffer(wb));
        }

        private static RangeTemplate Parse(IXLNamedRange range, TempSheetBuffer buff, RangeTemplate parent = null)
        {
            var prng = range.Ranges.First();
            var result = new RangeTemplate(range, buff,
                prng.RowCount(), prng.ColumnCount());

            var innerRanges = GetInnerRanges(prng).ToArray();

            var sheet = prng.Worksheet;

            for (int iRow = 1; iRow <= result._rowCnt; iRow++)
            {
                for (int iColumn = 1; iColumn <= result._colCnt; iColumn++)
                {
                    var xlCell = prng.Cell(iRow, iColumn);
                    if (innerRanges.Any(x => x.Ranges.Cells().Contains(xlCell)))
                        xlCell = null;
                    result._cells.Add(iRow, iColumn, xlCell);
                }
                if (iRow != result._rowCnt)
                    result._cells.AddNewRow();
            }

            result._mergedRanges = sheet.MergedRanges.Where(x => prng.Contains(x) && !innerRanges.Any(nr=>nr.Ranges.Any(r=>r.Contains(x)))).ToArray();
            sheet.MergedRanges.RemoveAll(result._mergedRanges.Contains);
            result._condFormats = sheet.ConditionalFormats
                .Where(f => prng.Contains(f.Range) && !innerRanges.Any(ir => ir.Ranges.Contains(f.Range)))
                .ToArray();
            if (result._rowCnt > 1)
            {
                // Exclude special row
                result._rowCnt--;

                result._rowRange = prng.Offset(0, 0, result._rowCnt, result._colCnt);
                result._optionsRow = prng.LastRow().Unsubscribed();
                result._optionsRowIsEmpty = !result._optionsRow.CellsUsed(false).Any();
                result._totalsCondFormats = sheet.ConditionalFormats
                    .Where(f => result._optionsRow.Contains(f.Range) && !innerRanges.Any(ir => ir.Ranges.Contains(f.Range)))
                    .ToArray();
                var rs = prng.RangeAddress.FirstAddress.RowNumber;
                result._condFormats = result._condFormats.Where(x => x.Range.RangeAddress.FirstAddress.RowNumber - rs + 1 <= result._rowCnt).ToArray();
            }else
                result._totalsCondFormats = new IXLConditionalFormat[0];

            result._subranges = innerRanges.Select(rng =>
            {
                var tpl = Parse(rng, buff, result);
                tpl._buff = result._buff;
                return tpl;
            }).ToArray();

            result.ParseTags(prng);

            if (result._rangeOption != null)
            {
                var source = result._rangeOption.GetParameter("source");
                if (!string.IsNullOrEmpty(source)) result.Source = source;
            }

            return result;
        }

        private static IEnumerable<IXLNamedRange> GetInnerRanges(IXLRange prng)
        {
            var containings = prng.GetContainingNames().ToArray();
            return from nr in containings
                let br = nr.Ranges
                    .Any(rng => containings
                        .Where(rr => rr != nr)
                        .SelectMany(rr => rr.Ranges)
                        .Any(r => r.Contains(rng)))
                where !br
                select nr;
        }

        public IReportBuffer Generate(object[] items)
        {
            var evaluator = new FormulaEvaluator();
            evaluator.AddVariable("items", items);
            _rangeTags.Reset();

            if (IsHorizontal)
            {
                HorizontalTable(items, evaluator);
            }
            else
            {
                VerticalTable(items, evaluator);
            }
            return _buff;
        }

        protected bool IsHorizontal
        {
            get { return (_rangeOption != null && _rangeOption.HasParameter("horizontal")) || (_rangeOption == null && _optionsRow == null); }
        }

        private void VerticalTable(object[] items, FormulaEvaluator evaluator)
        {
            var rangeStart = _buff.NextAddress;
            for (int i = 0; i < items.Length; i++)
            {
                var rowStart = _buff.NextAddress;
                IXLAddress rowEnd = null;
                int row = 1;
                var tags = _tags.CopyTo(_rowRange);

                // render row cells
                for (var iCell = 0; iCell < _cells.Count; iCell++)
                {
                    var cell = _cells[iCell];
                    if (cell.Row > _rowCnt)
                        break;

                    if (cell.CellType == TemplateCellType.None)
                    {
                        RenderSubrange(items[i], evaluator, cell, tags, ref iCell, ref row);
                    }
                    else if (cell.CellType == TemplateCellType.NewRow)
                    {
                        row++;
                        rowEnd = _buff.PrevAddress;
                        _buff.NewRow();
                        if (row == _rowCnt)
                            break;
                    }
                    else
                    {
                        RenderCell(items, i, evaluator, cell);
                    }
                }

                using (var newRowRng = _buff.GetRange(rowStart, rowEnd))
                {
                    foreach (var mrg in _mergedRanges.Where(r=>!_optionsRow.Contains(r)))
                    {
                        var newMrg = mrg.Relative(_rowRange, newRowRng).Unsubscribed();
                        newMrg.Merge(false);
                    }

                    if (_rowCnt > 1)
                    {
                        _buff.AddConditionalFormats(_condFormats, _rowRange, newRowRng);
                    }
                    tags.Execute(new ProcessingContext(newRowRng, items[i]));
                }
            }

            // Render options row
            if (!_optionsRowIsEmpty)
            {
                foreach (var cell in _cells.Where(c => c.Row == _rowCnt + 1).OrderBy(c => c.Column))
                {
                    RenderCell(evaluator, cell);
                }
                _buff.NewRow();
            }

            // Execute range options tags
            using (var resultRange = _buff.GetRange(rangeStart, _buff.PrevAddress))
            {
                if (_rowCnt == 1)
                {
                    var rows = resultRange.RowCount() - (_optionsRowIsEmpty ? 0 : 1);
                    _buff.AddConditionalFormats(_condFormats, _rowRange, resultRange.Offset(0, 0, rows, resultRange.ColumnCount()));
                }
                if (!_optionsRowIsEmpty)
                {
                    var optionsRow = resultRange.LastRow().AsRange().Unsubscribed();
                    foreach (var mrg in _mergedRanges.Where(r => _optionsRow.Contains(r)))
                    {
                        var newMrg = mrg.Relative(_optionsRow, optionsRow).Unsubscribed();
                        newMrg.Merge();
                    }
                    _buff.AddConditionalFormats(_totalsCondFormats, _optionsRow, optionsRow);
                }

                //Perhaps, this should be removed to support Pivot tags
                _rangeTags.Execute(new ProcessingContext(resultRange, new DataSource(items)));
            }
        }

        private void RenderCell(FormulaEvaluator evaluator, TemplateCell cell, params Parameter[] pars)
        {
            var value = cell.IsCalculated
                ? evaluator.Evaluate(cell.GetString(), pars)
                : cell.CellType == TemplateCellType.Formula ? cell.Formula : cell.Value;

            if (cell.CellType == TemplateCellType.Formula)
            {
                var r1c1 = cell.XLCell.GetFormulaR1C1(value.ToString());
                _buff.WriteFormulaR1C1(r1c1, cell.Style);
            }
            else
                _buff.WriteValue(value, cell.Style);
        }

        private void RenderCell(object[] items, int i, FormulaEvaluator evaluator, TemplateCell cell)
        {
            RenderCell(evaluator, cell, new Parameter("item", items[i]), new Parameter("index", i));
        }

        private void RenderSubrange(object item, FormulaEvaluator evaluator, TemplateCell cell, TagsList tags, ref int iCell, ref int row)
        {
            var start = _buff.NextAddress;
            // дочерний шаблон, к которому принадлежит ячейка
            var xlCell = _rowRange.Cell(cell.Row, cell.Column);
            var ownRng = _subranges.First(r => r._cells.Any(c => c.CellType != TemplateCellType.None && Equals(c.XLCell.Address, xlCell.Address)));
            var formula = "{{" + ownRng.Source.ReplaceLast("_", ".") + "}}";
            IEnumerable value = evaluator.Evaluate(formula, new Parameter(Name, item)) as IEnumerable;

            if (value != null)
            {
                var valArr = value.Cast<object>().ToArray();
                ownRng.Generate(valArr);

                if (ownRng.IsHorizontal)
                {
                    iCell += ownRng._colCnt - 1;
                    int shiftLen = ownRng._colCnt * (valArr.Length - 1);
                    tags.Where(tag => tag.Cell.Row == cell.Row && tag.Cell.Column > cell.Column)
                        .ForEach(t =>
                        {
                            t.Cell.Column += shiftLen;
                            t.Cell.XLCell = _rowRange.Cell(t.Cell.Row, t.Cell.Column);
                        });
                }
                else
                {
                    row += ownRng._rowCnt - 1;
                    while (_cells[iCell].Row <= row+1)
                        iCell++;

                    int shiftLen = ownRng._rowCnt * (valArr.Length - 1);
                    tags.Where(tag => tag.Cell.Row > cell.Row)
                        .ForEach(t =>
                        {
                            t.Cell.Row += shiftLen;
                            t.Cell.XLCell = _rowRange.Cell(t.Cell.Row, t.Cell.Column);
                        });
                }
            }
            var rng = _buff.GetRange(start, _buff.PrevAddress).Unsubscribed();
            var rangeName = ownRng.Name;
            var dnr = rng.Worksheet.Workbook.NamedRange(rangeName);
            dnr.SetRefersTo(rng);
        }

        private void HorizontalTable(object[] items, FormulaEvaluator evaluator)
        {
            var rangeStart = _buff.NextAddress;
            var tags = _tags.CopyTo(_rowRange);
            for (int i = 0; i < items.Length; i++)
            {
                var clmnStart = _buff.NextAddress;
                foreach (var cell in _cells)
                {
                    RenderCell(items, i, evaluator, cell);
                }
                using (var newClmnRng = _buff.GetRange(clmnStart, _buff.PrevAddress))
                    tags.Execute(new ProcessingContext(newClmnRng, items[i]));
            }
            /*using (var resultRange = _buff.GetRange(rangeStart, _buff.PrevAddress))
                _rangeTags.Execute(new ProcessingContext(resultRange, new DataSource(items)));*/
        }

        private void ParseTags(IXLRange range)
        {
            var innerRanges = range.GetContainingNames().ToArray();
            var cells = from c in _cells
                        let value = c.GetString()
                        where (value.StartsWith("<<") || value.EndsWith(">>"))
                            && !innerRanges.Any(nr => { using (var r = nr.Ranges) using (var cr = c.XLCell.AsRange()) return r.Contains(cr); })
                        select c;

            foreach (var cell in cells)
            {
                OptionTag[] tags;
                string newValue;
                if (cell.CellType == TemplateCellType.Formula)
                {
                    tags = _tagsEvaluator.Parse(cell.Formula, range, cell, out newValue);
                    cell.Formula = newValue;
                }
                else
                {
                    tags = _tagsEvaluator.Parse(cell.GetString(), range, cell, out newValue);
                    cell.Value = newValue;
                }
                if (cell.Row > _rowCnt)
                    _rangeTags.AddRange(tags);
                else
                    _tags.AddRange(tags);
            }

            _rangeOption = _rangeTags.GetAll<RangeOptionTag>().FirstOrDefault();
        }

        public void RangeTagsApply(IXLRange range, object[] items)
        {
            _rangeTags.Execute(new ProcessingContext(range, new DataSource(items)));
        }
    }
}