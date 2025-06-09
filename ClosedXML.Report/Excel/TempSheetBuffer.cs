using System;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Excel
{
    internal class TempSheetBuffer: IReportBuffer
    {
        private const string SheetName = "__temp_buffer";
        private readonly XLWorkbook _wb;
        private IXLWorksheet _sheet;
        private int _row;
        private int _clmn;
        private int _minRow;
        private int _minClmn;
        private int _prevrow;
        private int _prevclmn;

        public TempSheetBuffer(XLWorkbook wb)
        {
            _wb = wb;
            Init();
        }

        public IXLAddress NextAddress => _sheet.Cell(_row, _clmn).Address;
        public IXLAddress PrevAddress => _sheet.Cell(_prevrow, _prevclmn).Address;

        private void Init()
        {
            if (_sheet == null)
            {
                if (!_wb.TryGetWorksheet(SheetName, out _sheet))
                {
                    _sheet = _wb.AddWorksheet(SheetName);
                }
                _sheet.Visibility = XLWorksheetVisibility.VeryHidden;
            }
            _row = _minRow = _prevrow = 1;
            _clmn = _minClmn = _prevclmn = 1;
            Clear();
            _sheet.Style = _wb.Worksheets.First().Style;
        }

        public IXLCell WriteValue(object value, IXLCell settingCell)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            if (settingCell != null)
            {
                xlCell.CopyFrom(settingCell);
            }

            try
            {
                var cellValue = XLCellValueConverter.FromObject(value);
                xlCell.SetValue(cellValue);
            }
            catch (ArgumentException)
            {
                xlCell.SetValue(value?.ToString());
            }

            ChangeAddress(_row, _clmn + 1);
            return xlCell;
        }

        public IXLCell WriteFormulaR1C1(string formula, IXLCell settingCell)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            xlCell.CopyFrom(settingCell);
            xlCell.SetFormulaR1C1(formula);
            ChangeAddress(_row, _clmn + 1);
            return xlCell;
        }

        public void NewRow()
        {
            if (_clmn > 1)
                _clmn--;
            ChangeAddress(_row + 1, _minClmn);
            _minRow = _row;
        }

        public void NewRow(IXLAddress startAddr)
        {
            if (_clmn > 1)
                _clmn--;
            ChangeAddress(_row + 1, startAddr.ColumnNumber);
            _minRow = _row;
        }

        public void NewColumn(IXLAddress startAddr)
        {
            if (_clmn > 1)
                _clmn--;
            ChangeAddress(startAddr.RowNumber, _clmn + 1);
            _minClmn = _clmn;
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
            var tempRng = _sheet.Range(firstCell, LastCell);

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

            foreach (var picture in _sheet.Pictures)
            {
                var tgtPic = picture.CopyTo(tgtSheet);
                //var relAddress = picture.TopLeftCell.Relative(range.RangeAddress.FirstAddress);
                var tgtCell = range.RangeAddress.FirstAddress.Offset(picture.TopLeftCell.Address);
                tgtPic.MoveTo(tgtCell);
            }

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

        public IXLCell LastCell
        {
            get
            {
                var rowNumber = Math.Max(_prevrow, _sheet.RowsUsed().LastOrDefault()?.RowNumber() ?? 1);
                var columnNumber = Math.Max(_prevclmn, _sheet.ColumnsUsed().LastOrDefault()?.ColumnNumber() ?? 1);
                var lastCell = GetCell(rowNumber, columnNumber); //_sheet.Cell(_prevrow, _prevclmn);
                return lastCell;
            }
        }

        public void SetPrevCellToLastUsed()
        {
            var lastUsed = _sheet.LastCellUsed();
            var clmn = _clmn < lastUsed.Address.ColumnNumber
                ? lastUsed.Address.ColumnNumber + 1
                : _clmn;

            ChangeAddress(lastUsed.Address.RowNumber, clmn);
            NewRow();
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

        public void Dispose()
        {
            var namedRanges = _wb.DefinedNames
                .Where(nr => nr.Ranges.Any(r => r.Worksheet?.Name == SheetName))
                .ToList();
            namedRanges.ForEach(nr => nr.Delete());

            _wb.Worksheets.Delete(SheetName);
        }
    }
}
