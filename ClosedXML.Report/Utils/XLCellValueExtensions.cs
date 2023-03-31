using System;
using ClosedXML.Excel;

namespace ClosedXML.Report.Utils;

internal static class XLCellValueConverter
{
    public static XLCellValue FromObject(object obj, IFormatProvider provider = null)
    {
        return obj switch
        {
            null => Blank.Value,
            Blank blank => blank,
            bool logical => logical,
            string text => text,
            XLError error => error,
            DateTime dateTime => dateTime,
            TimeSpan timeSpan => timeSpan,
            sbyte number => number,
            byte number => number,
            short number => number,
            ushort number => number,
            int number => number,
            uint number => number,
            long number => number,
            ulong number => number,
            float number => number,
            double number => number,
            decimal number => number,
            _ => Convert.ToString(obj, provider)
        };
    }
}
