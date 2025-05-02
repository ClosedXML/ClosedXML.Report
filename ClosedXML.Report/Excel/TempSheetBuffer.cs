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
        private CellPosition _cellPosition;
        private CellPosition _minCellPosition;
        private CellPosition _prevCellPosition;
        private CellPosition _maxCellPosition;

        public TempSheetBuffer(XLWorkbook wb)
        {
            _wb = wb;
            Init();
        }

        public IXLAddress NextAddress => _sheet.Cell(_cellPosition.Row, _cellPosition.Column).Address;
        public IXLAddress PrevAddress => _sheet.Cell(_prevCellPosition.Row, _prevCellPosition.Column).Address;

        private void Init()
        {
            if (_sheet == null)
            {
                if (!_wb.TryGetWorksheet(SheetName, out _sheet)) 
                    _sheet = _wb.AddWorksheet(SheetName);
                
                _sheet.Visibility = XLWorksheetVisibility.VeryHidden;
            }

            _cellPosition = new CellPosition()
            {
                Row = _minCellPosition.Row = _prevCellPosition.Row = 1,
                Column = _minCellPosition.Column = _prevCellPosition.Column = 1
            };
            _maxCellPosition.Row = _maxCellPosition.Column = 1;
            
            Clear();
            _sheet.Style = _wb.Worksheets.First().Style;
        }

        public IXLCell WriteCellValue(object value, IXLCell settingCell)
        {
            var xlCell = _sheet.Cell(_cellPosition.Row, _cellPosition.Column);
            if (settingCell != null)
            {
                xlCell.CopyFrom(settingCell);
            }

            try
            {
                xlCell.SetValue(XLCellValueConverter.FromObject(value));
            }
            catch (ArgumentException)
            {
                xlCell.SetValue(value?.ToString());
            }

            UpdateMaxAddress(_cellPosition);
            ChangeAddress(_cellPosition.Row, _cellPosition.Column + 1);
            return xlCell;
        }
        
        private void UpdateMaxAddress(CellPosition cellPosition)
        {
            if (cellPosition.Row > _maxCellPosition.Row) _maxCellPosition.Row = cellPosition.Row;
            if (cellPosition.Column > _maxCellPosition.Column) _maxCellPosition.Column = cellPosition.Column;
        }

        public IXLCell WriteFormulaR1C1(string formula, IXLCell settingCell)
        {
            var xlCell = _sheet.Cell(_cellPosition.Row, _cellPosition.Column);
            xlCell.CopyFrom(settingCell);
            xlCell.SetFormulaR1C1(formula);
            UpdateMaxAddress(_cellPosition);
            ChangeAddress(_cellPosition.Row, _cellPosition.Column + 1);
            return xlCell;
        }

        public void NewRow()
        {
            StepBackColumn();
            ChangeAddress(_cellPosition.Row + 1, _minCellPosition.Column);
            _minCellPosition.Row = _cellPosition.Row;
        }
        
        private void StepBackColumn()
        {
            if (_cellPosition.Column > 1)
                _cellPosition.Column--;
        }

        public void NewRow(IXLAddress startAddr)
        {
            StepBackColumn();
            ChangeAddress(_cellPosition.Row + 1, startAddr.ColumnNumber);
            _minCellPosition.Row = _cellPosition.Row;
        }

        public void NewColumn(IXLAddress startAddr)
        {
            StepBackColumn();
            ChangeAddress(startAddr.RowNumber, _cellPosition.Column + 1);
            _minCellPosition.Column = _cellPosition.Column;
        }

        public IXLRange GetRange(IXLAddress startAddr, IXLAddress endAddr)
        {
            return _sheet.Range(startAddr, endAddr);
        }

        public IXLCell GetCell(CellPosition cellPosition)
        {
            return _sheet.Cell(cellPosition.Row, cellPosition.Column);
        }

        private void ChangeAddress(int row, int column)
        {
            _prevCellPosition = _cellPosition;
            _cellPosition = new CellPosition{
                Row = row,
                Column = column
            };
        }

        public IXLRange CopyTo(IXLRange range)
        {
            var firstCell = _sheet.Cell(1, 1);
            var tempRng = _sheet.Range(firstCell, LastCellUsed);

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

       public IXLCell LastCellUsed => GetCell(_maxCellPosition);

        public void SetPrevCellToLastUsed()
        {
            var lastUsed = _sheet.LastCellUsed();
            var clmn = _cellPosition.Column < lastUsed.Address.ColumnNumber
                ? lastUsed.Address.ColumnNumber + 1
                : _cellPosition.Column;

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
