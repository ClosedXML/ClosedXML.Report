using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core.Exceptions;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Utils;
using MoreLinq;

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
        private readonly TemplateErrors _errors;
        private readonly FormulaEvaluator _evaluator;
        private RangeOptionTag _rangeOption;
        private TempSheetBuffer _buff;
        private IXLRange _rowRange;
        private IXLRangeRow _optionsRow;
        private bool _optionsRowIsEmpty = true;
        private bool _isSubrange;
        private IDictionary<string, object> _globalVariables;

        public string Source { get; private set; }
        public string Name { get; }

        internal RangeTemplate(string name, IXLRange range, TempSheetBuffer buff, TemplateErrors errors, IDictionary<string, object> globalVariables)
        {
            _rowRange = range;
            _cells = new TemplateCells(this);
            _tagsEvaluator = new TagsEvaluator();
            var wb = _rowRange.Worksheet.Workbook;
            _buff = buff;
            _errors = errors;
            _globalVariables = globalVariables;
            _tags = new TagsList(_errors);
            _rangeTags = new TagsList(_errors);
            Name = name;
            Source = name;
            var rangeName = name + "_tpl";
            if (wb.NamedRanges.TryGetValue(rangeName, out var namedRange))
            {
                namedRange.Add(range);
            }
            else
            {
                wb.NamedRanges.Add(rangeName, range);
            }

            _evaluator = new FormulaEvaluator();
        }

        internal RangeTemplate(string name, IXLRange range, TempSheetBuffer buff, int rowCnt, int colCnt, TemplateErrors errors, IDictionary<string, object> globalVariables) : this(name, range, buff, errors, globalVariables)
        {
            _rowCnt = rowCnt;
            _colCnt = colCnt;
        }


        public static RangeTemplate Parse(string name, IXLRange range, TemplateErrors errors, IDictionary<string, object> globalVariables)
        {
            var wb = range.Worksheet.Workbook;
            return Parse(name, range, new TempSheetBuffer(wb), errors, globalVariables);
        }

        private static RangeTemplate Parse(string name, IXLRange range, TempSheetBuffer buff, TemplateErrors errors, IDictionary<string, object> globalVariables)
        {
            var result = new RangeTemplate(name, range, buff,
                range.RowCount(), range.ColumnCount(), errors, globalVariables);

            var innerRanges = GetInnerRanges(range).ToArray();

            var sheet = range.Worksheet;

            for (int iRow = 1; iRow <= result._rowCnt; iRow++)
            {
                for (int iColumn = 1; iColumn <= result._colCnt; iColumn++)
                {
                    var xlCell = range.Cell(iRow, iColumn);
                    if (innerRanges.Any(x => x.Ranges.Cells().Contains(xlCell)))
                        xlCell = null;
                    result._cells.Add(iRow, iColumn, xlCell);
                }
                if (iRow != result._rowCnt)
                    result._cells.AddNewRow();
            }

            result._mergedRanges = sheet.MergedRanges.Where(x => range.Contains(x) && !innerRanges.Any(nr=>nr.Ranges.Any(r=>r.Contains(x)))).ToArray();
            sheet.MergedRanges.RemoveAll(result._mergedRanges.Contains);

            result.ParseTags(range);

            if (result._rowCnt > 1 && !result.IsHorizontal)
            {
                // Exclude special row
                result._rowCnt--;

                result._rowRange = range.Offset(0, 0, result._rowCnt, result._colCnt);
                result._optionsRow = range.LastRow();
                result._optionsRowIsEmpty = !result._optionsRow.CellsUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges).Any();
            }

            result._subranges = innerRanges.SelectMany(nrng => nrng.Ranges,
                (nr, rng) =>
                {
                    var tpl = Parse(nr.Name, rng, buff, errors, globalVariables);
                    tpl._buff = result._buff;
                    tpl._isSubrange = true;
                    tpl._globalVariables = globalVariables;
                    return tpl;
                }).ToArray();

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
            _evaluator.AddVariable("items", items);
            foreach (var v in _globalVariables)
            {
                _evaluator.AddVariable("@"+v.Key, v.Value);
            }
            _rangeTags.Reset();

            if (IsHorizontal)
            {
                HorizontalTable(items, _evaluator);
            }
            else
            {
                VerticalTable(items, _evaluator);
            }
            return _buff;
        }

        protected bool IsHorizontal => (_rangeOption != null && _rangeOption.HasParameter("horizontal"))
                                       || (_rowCnt == 1 && _optionsRow == null)
                                       || _colCnt == 1;

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
                        if (row > _rowCnt)
                            break;
                    }
                    else
                    {
                        RenderCell(items, i, evaluator, cell);
                    }
                }

                var newRowRng = _buff.GetRange(rowStart, rowEnd);
                foreach (var mrg in _mergedRanges.Where(r=>!_optionsRow.Contains(r)))
                {
                    var newMrg = mrg.Relative(_rowRange, newRowRng);
                    newMrg.Merge(false);
                }

                tags.Execute(new ProcessingContext(newRowRng, items[i], evaluator));
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
            var resultRange = _buff.GetRange(rangeStart, _buff.PrevAddress);
            if (!_optionsRowIsEmpty)
            {
                var optionsRow = resultRange.LastRow().AsRange();
                foreach (var mrg in _mergedRanges.Where(r => _optionsRow.Contains(r)))
                {
                    var newMrg = mrg.Relative(_optionsRow, optionsRow);
                    newMrg.Merge();
                }
            }

            if (_isSubrange)
            {
                _rangeTags.Execute(new ProcessingContext(resultRange, new DataSource(items), evaluator));
                // if the range was increased by processing tags (for example, Group), move the buffer to the last cell
                _buff.SetPrevCellToLastUsed(); 
            }
        }

        private void RenderCell(FormulaEvaluator evaluator, TemplateCell cell, params Parameter[] pars)
        {
            if (cell.CellType != TemplateCellType.Formula && cell.CellType != TemplateCellType.Value)
            {
                _buff.WriteValue(null, null);
                return;
            }

            object value;
            try
            {
                value = cell.IsCalculated
                    ? evaluator.Evaluate(cell.GetString(), pars)
                    : cell.CellType == TemplateCellType.Formula ? cell.Formula : cell.Value;
            }
            catch (ParseException ex)
            {
                _buff.WriteValue(ex.Message, cell.XLCell);
                _buff.GetCell(_buff.PrevAddress.RowNumber, _buff.PrevAddress.ColumnNumber).Style.Font.FontColor = XLColor.Red;
                _errors.Add(new TemplateError(ex.Message, cell.XLCell.AsRange()));
                return;
            }

            IXLCell xlCell;
            if (cell.CellType == TemplateCellType.Formula)
            {
                var r1c1 = cell.XLCell.GetFormulaR1C1(value.ToString());
                xlCell = _buff.WriteFormulaR1C1(r1c1, cell.XLCell);
            }
            else
            {
                xlCell = _buff.WriteValue(value, cell.XLCell);
            }

            string EvalString(string str)
            {
                try
                {
                    return evaluator.Evaluate(str, pars).ToString();
                }
                catch (ParseException ex)
                {
                    _errors.Add(new TemplateError(ex.Message, cell.XLCell.AsRange()));
                    return ex.Message;
                }
            }

            if (xlCell.HasComment)
            {
                var comment = EvalString(xlCell.Comment.Text);
                xlCell.Comment.ClearText();
                xlCell.Comment.AddText(comment);
            }

            if (xlCell.HasHyperlink)
            {
                if (xlCell.Hyperlink.IsExternal)
                    xlCell.Hyperlink.ExternalAddress = new Uri(EvalString(xlCell.Hyperlink.ExternalAddress.ToString()));
                else
                    xlCell.Hyperlink.InternalAddress = EvalString(xlCell.Hyperlink.InternalAddress);
            }

            if (xlCell.HasRichText)
            {
                var richText = EvalString(xlCell.RichText.Text);
                xlCell.RichText.ClearText();
                xlCell.RichText.AddText(richText);
            }
        }



        private void RenderCell(object[] items, int i, FormulaEvaluator evaluator, TemplateCell cell)
        {
            RenderCell(evaluator, cell, new Parameter("item", items[i]), new Parameter("index", i));
        }

        private void RenderSubrange(object item, FormulaEvaluator evaluator, TemplateCell cell, TagsList tags, ref int iCell, ref int row)
        {
            var start = _buff.NextAddress;
            // the child template to which the cell belongs
            var xlCell = _rowRange.Cell(cell.Row, cell.Column);
            var ownRng = _subranges.First(r => r._cells.Any(c => c.CellType != TemplateCellType.None && c.XLCell != null && Equals(c.XLCell.Address, xlCell.Address)));
            var formula = ownRng.Source.ReplaceLast("_", ".");

            if (evaluator.Evaluate(formula, new Parameter(Name, item)) is IEnumerable value)
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
                    // move current template cell to next (skip subrange)
                    row += ownRng._rowCnt+1;
                    while (_cells[iCell].Row <= row-1)
                        iCell++;

                    iCell--; // roll back. After it became clear that it was too much, we must go back.

                    int shiftLen = ownRng._rowCnt * (valArr.Length - 1);
                    tags.Where(tag => tag.Cell.Row > cell.Row)
                        .ForEach(t =>
                        {
                            t.Cell.Row += shiftLen;
                            t.Cell.XLCell = _rowRange.Cell(t.Cell.Row, t.Cell.Column);
                        });
                }
            }

            var rng = _buff.GetRange(start, _buff.PrevAddress);
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
                    if (cell.CellType == TemplateCellType.None)
                        throw new NotSupportedException("Horizontal range does not support subranges.");
                    else if (cell.CellType != TemplateCellType.NewRow)
                        RenderCell(items, i, evaluator, cell);
                    else
                        _buff.NewRow();
                }

                var newClmnRng = _buff.GetRange(clmnStart, _buff.PrevAddress);
                foreach (var mrg in _mergedRanges.Where(r => _optionsRow == null || !_optionsRow.Contains(r)))
                {
                    var newMrg = mrg.Relative(_rowRange, newClmnRng);
                    newMrg.Merge(false);
                }

                tags.Execute(new ProcessingContext(newClmnRng, items[i], evaluator));

                if (_rowCnt > 1)
                    _buff.NewColumn();
            }

            var worksheet = _rowRange.Worksheet;
            var colNumbers = _cells.Where(xc => xc.XLCell != null)
                .Select(xc => xc.XLCell.Address.ColumnNumber)
                .Distinct()
                .ToArray();
            var widths = colNumbers
                .Select(c => worksheet.Column(c).Width)
                .ToArray();
            var firstCol = colNumbers.Min();
            foreach (var col in Enumerable.Range(rangeStart.ColumnNumber, _buff.PrevAddress.ColumnNumber))
            {
                worksheet.Column(firstCol + col - 1).Width = widths[(col - 1) % widths.Length];
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
                            && !innerRanges.Any(nr => nr.Ranges.Contains(c.XLCell.AsRange()))
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
                if (cell.Row == _rowCnt)
                    _rangeTags.AddRange(tags);
                else
                    _tags.AddRange(tags);
            }

            _rangeOption = _rangeTags.GetAll<RangeOptionTag>().Union(_tags.GetAll<RangeOptionTag>()).FirstOrDefault();
        }

        public void RangeTagsApply(IXLRange range, object[] items)
        {
            _rangeTags.Execute(new ProcessingContext(range, new DataSource(items), _evaluator));
        }
    }
}
