using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.Excel
{
    public static class XlExtensions
    {
        public static readonly Regex A1SimpleRegex = new Regex(
            //  @"(?<=\W)" // Start with non word
            @"(?<Reference>" // Start Group to pick
            + @"(?<Sheet>" // Start Sheet Name, optional
            + @"("
            + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
            // Sheet name with special characters, surrounding apostrophes are required
            + @"|"
            + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
            + @")"
            + @"!)?" // End Sheet Name, optional
            + @"(?<Range>" // Start range
            + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
            + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
            + @"|"
            + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
            + @"|"
            + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
            + @")" // End Range
            + @")" // End Group to pick
            //+ @"(?=\W)" // End with non word
            , RegexOptions.Compiled);

        /// <summary>
        /// Find ranges within which contains the specified range (completely).
        /// </summary>
        /// <param name="range">range</param>
        public static IEnumerable<IXLNamedRange> GetContainerNames(this IXLRange range)
        {
            return range.Worksheet.Workbook.NamedRanges.Where(x => x.Ranges.Where(r => !r.Equals(range)).Any(r => r.Contains(range)));
        }

        public static bool Contains(this IXLRangeAddress rangeAddress, IXLAddress address)
        {
            return rangeAddress.FirstAddress.RowNumber <= address.RowNumber &&
                   address.RowNumber <= rangeAddress.LastAddress.RowNumber &&
                   rangeAddress.FirstAddress.ColumnNumber <= address.ColumnNumber &&
                   address.ColumnNumber <= rangeAddress.LastAddress.ColumnNumber;
        }

        public static bool Contains(this IXLRangeAddress rangeAddress, IXLRangeAddress address)
        {
            return rangeAddress.Contains(address.FirstAddress) && rangeAddress.Contains(address.LastAddress);
        }

        internal static void ShiftRows(this IXLRangeBase range, int rowCount)
        {
            var firstAddress = range.RangeAddress.FirstAddress;
            var lastAddress = range.RangeAddress.LastAddress;
            range.RangeAddress.FirstAddress = range.Worksheet.Cell(firstAddress.RowNumber + rowCount, firstAddress.ColumnNumber).Address;
            range.RangeAddress.LastAddress = range.Worksheet.Cell(lastAddress.RowNumber + rowCount, lastAddress.ColumnNumber).Address;
        }

        internal static void ExtendRows(this IXLRangeBase range, int rowCount, bool down = true)
        {
            if (down)
            {
                var lastAddress = range.RangeAddress.LastAddress;
                range.RangeAddress.LastAddress = range.Worksheet.Cell(lastAddress.RowNumber + rowCount, lastAddress.ColumnNumber).Address;
            }
            else
            {
                var firstAddress = range.RangeAddress.FirstAddress;
                range.RangeAddress.FirstAddress = range.Worksheet.Cell(firstAddress.RowNumber - rowCount, firstAddress.ColumnNumber).Address;
            }
        }

        internal static KeyValuePair<string, IXLRangeAddress>[] GetRangeParameters(this IXLWorksheet ws, string formulaA1)
        {
            if (formulaA1.IsNullOrWhiteSpace()) return null;

            var regex = A1SimpleRegex;
            List<KeyValuePair<string, IXLRangeAddress>> result = new List<KeyValuePair<string, IXLRangeAddress>>();

            foreach (var match in regex.Matches(formulaA1).Cast<Match>())
            {
                var matchValue = match.Value;
                
                if (matchValue.Contains('!'))
                {
                    var split = matchValue.Split('!');
                    var first = split[0];
                    var wsName = first.StartsWith("'") ? first.Substring(1, first.Length - 2) : first;
                    matchValue = split[1];

                    IXLWorksheet refWs;
                    if (ws.Workbook.Worksheets.TryGetWorksheet(wsName, out refWs))
                    {
                        ws = refWs;
                    }
                }
                result.Add(new KeyValuePair<string, IXLRangeAddress>(matchValue, ws.Range(matchValue).RangeAddress));
            }
            return result.ToArray();
        }

        /// <summary>
        /// Get the named ranges that contains the specified range (completely).
        /// </summary>
        /// <param name="range">range</param>
        public static IEnumerable<IXLNamedRange> GetContainingNames(this IXLRange range)
        {
            return range.Worksheet.Workbook.NamedRanges.Where(x => x.Ranges.Where(r => r.Worksheet.Position == range.Worksheet.Position
                                                                                       && !r.Equals(range)).Any(range.Contains));
        }

        /// <summary>
        /// Get range coordinates relative to another range.
        /// </summary>
        /// <param name="range">range</param>
        /// <param name="baseAddr">Reference system. Coordinates are calculated relative to this range.</param>
        public static IXLRangeAddress Relative(this IXLRangeAddress range, IXLRangeAddress baseAddr)
        {
            using (var xlRange = baseAddr.Worksheet.Range(
                range.FirstAddress.RowNumber - baseAddr.FirstAddress.RowNumber + 1,
                range.FirstAddress.ColumnNumber - baseAddr.FirstAddress.ColumnNumber + 1,
                range.LastAddress.RowNumber - baseAddr.FirstAddress.RowNumber + 1,
                range.LastAddress.ColumnNumber - baseAddr.FirstAddress.ColumnNumber + 1))
                return xlRange.RangeAddress;
        }

        /// <summary>
        /// Get range coordinates relative to another range.
        /// </summary>
        /// <param name="cell">range</param>
        /// <param name="baseAddr">Reference system. Coordinates are calculated relative to this range.</param>
        public static IXLAddress Relative(this IXLCell cell, IXLAddress baseAddr)
        {
            return baseAddr.Worksheet.Cell(
                cell.Address.RowNumber - baseAddr.RowNumber + 1,
                cell.Address.ColumnNumber - baseAddr.ColumnNumber + 1).Address;
        }


        /// <summary>
        /// Get range relative to another range.
        /// </summary>
        /// <param name="range">range</param>
        /// <param name="baseRange">Coordinate system. Coordinates are calculated relative to this range.</param>
        /// <param name="targetBase"></param>
        public static IXLRange Relative(this IXLRangeBase range, IXLRangeBase baseRange, IXLRangeBase targetBase)
        {
            using (var xlRange = targetBase.Worksheet.Range(
                range.RangeAddress.FirstAddress.RowNumber - baseRange.RangeAddress.FirstAddress.RowNumber + 1,
                range.RangeAddress.FirstAddress.ColumnNumber - baseRange.RangeAddress.FirstAddress.ColumnNumber + 1,
                range.RangeAddress.LastAddress.RowNumber - baseRange.RangeAddress.FirstAddress.RowNumber + 1,
                range.RangeAddress.LastAddress.ColumnNumber - baseRange.RangeAddress.FirstAddress.ColumnNumber + 1))
            {
                var type = targetBase.GetType();
                var method = type.GetMethod("Range", new[] {typeof(IXLRangeAddress)});
                return (IXLRange)method.Invoke(targetBase, new object[] { xlRange.RangeAddress });
                //return targetBase.Range(xlRange.RangeAddress);
            }
        }
        public static void Subtotal(this IXLRange range, int groupBy, string function, int[] totalList, bool replace = true, bool pageBreaks = false, bool summaryAbove = false)
        {
            using (var subtotal = new Subtotal(range, summaryAbove))
            {
                if (replace)
                    subtotal.Unsubtotal();
                var summaries = totalList.Select(x => new SubtotalSummaryFunc(function, x)).ToArray();
                subtotal.AddGrandTotal(summaries);
                subtotal.GroupBy(groupBy, summaries, pageBreaks);
            }
        }

        public static bool IsSummary(this IXLRangeRow row)
        {
            return row.Cells(x => x.HasFormula && x.FormulaA1.ToLower().StartsWith("subtotal(")).Any();
        }

        public static bool IsEmpty(this IXLRangeRow row)
        {
            return !row.Cells(x => x.HasFormula || !x.GetString().IsNullOrWhiteSpace()).Any();
        }

        public static T Unsubscribed<T>(this T range) where T : IXLRangeBase
        {
            range.Dispose();
            return range;
        }

        public static void CopyStylesFrom(this IXLRangeBase trgtRow, IXLRangeBase srcRow)
        {
            trgtRow.Style = srcRow.Style;
            var srcCells = srcRow.Cells(true, true).ToArray();
            for (int i = 0; i < srcCells.Length; i++)
            {
                var rela = srcCells[i].Relative(srcRow.RangeAddress.FirstAddress);
                var trgtCell = trgtRow.RangeAddress.FirstAddress.Offset(rela);
                trgtCell.Style = srcCells[i].Style;
                //trgtCells[i].Style = srcCells[i].Style;
            }
            //trgtRow.CopyConditionalFormatsFrom(srcRow);
        }

        public static void CopyFrom(this IXLConditionalFormat targetFormat, IXLConditionalFormat srcFormat)
        {
            var type = targetFormat.GetType();
            var method = type.GetMethod("CopyFrom", BindingFlags.Instance | BindingFlags.Public);
            method.Invoke(targetFormat, new object[] { srcFormat });
        }

        public static void SuspendEvents(this IXLWorksheet sheet)
        {
            var type = sheet.GetType();
            var method = type.GetMethod("SuspendEvents", BindingFlags.Instance | BindingFlags.Public);
            method.Invoke(sheet, new object[0]);
        }

        public static void ResumeEvents(this IXLWorksheet sheet)
        {
            var type = sheet.GetType();
            var method = type.GetMethod("ResumeEvents", BindingFlags.Instance | BindingFlags.Public);
            method.Invoke(sheet, new object[0]);
        }

        public static int RowCount(this IXLRangeAddress address)
        {
            return address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        }

        internal static string GetFormulaR1C1(this IXLCell cell, string value)
        {
            var type = cell.GetType();
            var method = type.GetMethod("GetFormulaR1C1", BindingFlags.Instance | BindingFlags.NonPublic);
            return (string)method.Invoke(cell, new object[] { value });
        }

        internal static string GetFormulaA1(this IXLCell cell, string value)
        {
            var type = cell.GetType();
            var method = type.GetMethod("GetFormulaA1", BindingFlags.Instance | BindingFlags.NonPublic);
            return (string)method.Invoke(cell, new object[] { value });
        }

        internal static void CopyRelative(this IXLConditionalFormat format, IXLRangeBase fromRange, IXLRangeBase toRange, bool expand)
        {
            var frmtRng = format.Range.Relative(fromRange, toRange);
            if (expand && toRange.RangeAddress.RowCount() != format.Range.RowCount())
                frmtRng = frmtRng.Offset(0, 0, toRange.RangeAddress.RowCount(), frmtRng.ColumnCount()).Unsubscribed();
            var newFrmt = frmtRng.AddConditionalFormat();
            newFrmt.CopyFrom(format);
        }

        internal static void CopyConditionalFormatsFrom(this IXLRangeBase targetRange, IXLRangeBase srcRange)
        {
            var sheet = targetRange.Worksheet;
            sheet.SuspendEvents();
            foreach (var conditionalFormat in sheet.ConditionalFormats.Where(c => c.Range.Intersects(srcRange)).ToList())
            {
                conditionalFormat.CopyRelative(srcRange, targetRange, false);
            }
            sheet.ResumeEvents();
        }

        public static IXLRange Offset(this IXLRange range, int rowsOffset, int columnOffset)
        {
            return Offset(range, rowsOffset, columnOffset, range.RowCount(), range.ColumnCount());
        }

        public static IXLRange Offset(this IXLRange range, int rowsOffset, int columnOffset, int rows, int columns)
        {
            var sheet = range.Worksheet;
            return sheet.Range(
                range.RangeAddress.FirstAddress.RowNumber + rowsOffset,
                range.RangeAddress.FirstAddress.ColumnNumber + columnOffset,
                range.RangeAddress.FirstAddress.RowNumber + rowsOffset + rows - 1,
                range.RangeAddress.FirstAddress.ColumnNumber + columnOffset + columns - 1);
        }

        public static IXLCell Offset(this IXLAddress addr, IXLAddress offset)
        {
            var sheet = addr.Worksheet;
            return sheet.Cell(
                addr.RowNumber + offset.RowNumber - 1,
                addr.ColumnNumber + offset.ColumnNumber - 1);
        }

        public static void SetCalcEngineCacheExpressions(this IXLWorksheet worksheet, bool value)
        {
            var wsType = worksheet.GetType();
            var calcEngine = wsType.GetProperty("CalcEngine", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(worksheet);
            calcEngine.GetType().GetProperty("CacheExpressions").SetValue(calcEngine, value);
        }

        /* ClosedXML issue #686 */
        public static void ReplaceCFFormulaeToR1C1(this IXLWorksheet worksheet)
        {
            foreach (var format in worksheet.ConditionalFormats)
            {
                var source = format.Range.FirstCell();
                foreach (var v in format.Values.Where(v => v.Value.IsFormula).ToList())
                {
                    var f = v.Value.Value;
                    var r1c1 = source.GetFormulaR1C1(f);
                    format.Values[v.Key] = new XLFormula("&=" + r1c1);
                }
            }
        }

        /* ClosedXML issue #686 */
        public static void ReplaceCFFormulaeToA1(this IXLWorksheet worksheet)
        {
            foreach (var format in worksheet.ConditionalFormats)
            {
                var target = format.Range.FirstCell();
                foreach (var v in format.Values.Where(v => v.Value.Value.StartsWith("&=")).ToList())
                {
                    var f = v.Value.Value.Substring(1);
                    var a1 = target.GetFormulaA1(f);
                    format.Values[v.Key] = new XLFormula(a1);
                }
            }
        }

    }

    public enum XlCopyType
    {
        All = -4104,	// Everything will be pasted.
        AllExceptBorders = 7,	// Everything except borders will be pasted.
        AllMergingConditionalFormats = 14,	// Everything will be pasted and conditional formats will be merged.
        AllUsingSourceTheme = 13,	// Everything will be pasted using the source theme.
        ColumnWidths = 8,	// Copied column width is pasted.
        Comments = -4144,	// Comments are pasted.
        Formats = -4122,	// Copied source format is pasted.
        Formulas = -4123,	// Formulas are pasted.
        FormulasAndNumberFormats = 11,	// Formulas and Number formats are pasted.
        Values = -4163,	// Values are pasted.
        ValuesAndNumberFormats = 12	// Values and Number formats are pasted.
    }
}