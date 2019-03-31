using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Report.Excel
{
    internal class TempSheetBuffer: IReportBuffer
    {
        private const string SheetName = "__temp_buffer";
        private readonly XLWorkbook _wb;
        private IXLWorksheet _sheet;
        private int _row;
        private int _clmn;
        private int _prevrow;
        private int _prevclmn;

        public TempSheetBuffer(XLWorkbook wb)
        {
            _wb = wb;
            Init();
        }

        public IXLAddress NextAddress { get { return _sheet.Cell(_row, _clmn).Address; } }
        public IXLAddress PrevAddress { get { return _sheet.Cell(_prevrow, _prevclmn).Address; } }

        private void Init()
        {
            if (_sheet == null)
            {
                if (!_wb.TryGetWorksheet(SheetName, out _sheet))
                {
                    _sheet = _wb.AddWorksheet(SheetName);
                    _sheet.SetCalcEngineCacheExpressions(false);
                }
                _sheet.Visibility = XLWorksheetVisibility.VeryHidden;
            }
            _row = _prevrow = 1;
            _clmn = _prevclmn = 1;
            Clear();
            _sheet.Style = _wb.Worksheets.First().Style;
        }

        public void WriteValue(object value, IXLStyle cellStyle)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            try
            {
                xlCell.SetValue(value);
            }
            catch (ArgumentException)
            {
                xlCell.SetValue(value?.ToString());
            }

            xlCell.Style = cellStyle ?? _wb.Style;
            ChangeAddress(_row, _clmn + 1);
        }

        public void WriteFormulaR1C1(string formula, IXLStyle cellStyle)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            xlCell.Style = cellStyle;
            xlCell.SetFormulaR1C1(formula);
            ChangeAddress(_row, _clmn + 1);
        }

        public void NewRow()
        {
            if (_clmn > 1)
                _clmn--;
            ChangeAddress(_row + 1, 1);
        }

        public IXLRange GetRange(IXLAddress startAddr, IXLAddress endAddr)
        {
            return _sheet.Range(startAddr, endAddr);
        }

        public IXLCell GetCell(int row, int column)
        {
            return _sheet.Cell(row, column);
        }

        private void ChangeAddress(int row, int clmn)
        {
            _prevrow = _row;
            _prevclmn = _clmn;
            _row = row;
            _clmn = clmn;
        }

        public IXLRange CopyTo(IXLRange range)
        {
            var firstCell = _sheet.Cell(1, 1);
            var lastCell = _sheet.LastCellUsed(XLCellsUsedOptions.All) ?? firstCell;
            var tempRng = _sheet.Range(firstCell, lastCell);

            var rowDiff = tempRng.RowCount() - range.RowCount();
            if (rowDiff > 0)
                range.LastRow().RowAbove().InsertRowsBelow(rowDiff, true);
            else if (rowDiff < 0)
                range.Worksheet.Range(
                    range.LastRow().RowNumber() + rowDiff + 1,
                    range.FirstColumn().ColumnNumber(),
                    range.LastRow().RowNumber(),
                    range.LastColumn().ColumnNumber())
                .Delete(XLShiftDeletedCells.ShiftCellsUp);

            range.Worksheet.ConditionalFormats.Remove(c => c.Range.Intersects(range));

            var columnDiff = tempRng.ColumnCount() - range.ColumnCount();
            if (columnDiff > 0)
                range.InsertColumnsAfter(columnDiff, true);
            else if (columnDiff < 0)
                range.Worksheet.Range(
                    range.FirstRow().RowNumber(),
                    range.LastColumn().ColumnNumber() + columnDiff + 1,
                    range.LastRow().RowNumber(),
                    range.LastColumn().ColumnNumber())
                .Delete(XLShiftDeletedCells.ShiftCellsLeft);

            tempRng.CopyTo(range.FirstCell());

            var tgtSheet = range.Worksheet;
            var tgtStartRow = range.RangeAddress.FirstAddress.RowNumber;
            var srcRows = _sheet.Rows(tempRng.RangeAddress.FirstAddress.RowNumber, tempRng.RangeAddress.LastAddress.RowNumber);
            foreach (var row in srcRows)
            {
                var xlRow = tgtSheet.Row(row.RowNumber() + tgtStartRow-1);
                xlRow.OutlineLevel = row.OutlineLevel;
                if (row.IsHidden)
                    xlRow.Collapse();
                else
                    xlRow.Expand();
            }
            return range;
        }

        public void Clear()
        {
            var srcRows = _sheet.RowsUsed(XLCellsUsedOptions.All);
            foreach (var row in srcRows)
            {
                row.OutlineLevel = 0;
            }
            _sheet.Clear();
        }

        public void AddConditionalFormats(IEnumerable<IXLConditionalFormat> formats, IXLRangeBase fromRange, IXLRangeBase toRange)
        {
            foreach (var format in formats)
            {
                format.CopyRelative(fromRange, toRange, true);
            }
        }

        public void Dispose()
        {
            var namedRanges = _wb.NamedRanges
                .Where(nr => nr.Ranges.Any(r => r.Worksheet?.Name == SheetName))
                .ToList();
            namedRanges.ForEach(nr => nr.Delete());

            _wb.Worksheets.Delete(SheetName);
        }
    }
}
