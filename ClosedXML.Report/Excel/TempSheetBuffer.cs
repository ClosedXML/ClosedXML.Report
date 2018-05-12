using System;
using System.Collections.Generic;
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
        private int _prevrow;
        private int _prevclmn;
        private int _maxClmn;
        private int _maxRow;

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
                //_sheet.Visibility = XLWorksheetVisibility.VeryHidden;
            }
            _row = 1;
            _clmn = 1;
            _maxRow = _prevrow = 1;
            _maxClmn = _prevclmn = 1;
            Clear();
        }

        public void WriteValue(object value, IXLStyle cellStyle)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            xlCell.SetValue(value);
            xlCell.Style = cellStyle ?? _wb.Style;
            _maxClmn = Math.Max(_maxClmn, _clmn);
            _maxRow = Math.Max(_maxRow, _row);
            ChangeAddress(_row, _clmn + 1);
        }

        public void WriteFormulaR1C1(string formula, IXLStyle cellStyle)
        {
            var xlCell = _sheet.Cell(_row, _clmn);
            xlCell.Style = cellStyle;
            xlCell.SetFormulaR1C1(formula);
            _maxClmn = Math.Max(_maxClmn, _clmn);
            _maxRow = Math.Max(_maxRow, _row);
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
            var tempRng = _sheet.Range(_sheet.Cell(1, 1), _sheet.LastCellUsed()); //_sheet.Cell(_prevrow, _prevclmn));

            range.InsertRowsBelow(tempRng.RowCount() - range.RowCount(), true);
            range.InsertColumnsAfter(tempRng.ColumnCount() - range.ColumnCount(), true);
            tempRng.CopyTo(range.FirstCell());

            var tgtSheet = range.Worksheet;
            var tgtStartRow = range.RangeAddress.FirstAddress.RowNumber;
            using (var srcRows = _sheet.Rows(tempRng.RangeAddress.FirstAddress.RowNumber, tempRng.RangeAddress.LastAddress.RowNumber))
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
            using (var srcRows = _sheet.RowsUsed(true))
                foreach (var row in srcRows)
                {
                    row.OutlineLevel = 0;
                }
            _sheet.Clear();
        }

        public void AddConditionalFormats(IEnumerable<IXLConditionalFormat> formats, IXLRangeBase fromRange, IXLRangeBase toRange)
        {
            //var tempRng = _sheet.Range(_sheet.Cell(1, 1), _sheet.Cell(_prevrow, _prevclmn));
            foreach (var format in formats)
            {
                format.CopyRelative(fromRange, toRange, true);
            }
        }

        public void Dispose()
        {
            _wb.Worksheets.Delete(SheetName);
        }
    }
}