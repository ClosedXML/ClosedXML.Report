using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Report.Utils;

public static class IXLCellExtensions
{
    public static void SetObjectValue(this IXLCell cell, object value)
    {
        if (value == null)
        {
            cell.Value = Blank.Value;
        }
        else if (value is IEnumerable data && value is not string)
        {
            cell.InsertData(data);
        }
        else if (value is bool bv)
        {
            cell.Value = bv;
        }
        else if (value is string sv)
        {
            cell.Value = sv;
        }
        else if (value is XLError err)
        {
            cell.Value = err;
        }
        else if (value is DateTime dtv)
        {
            cell.Value = dtv;
        }
        else if (value is TimeSpan tsv)
        {
            cell.Value = tsv;
        }
        else if (value is sbyte sbv)
        {
            cell.Value = sbv;
        }
        else if (value is byte btv)
        {
            cell.Value = btv;
        }
        else if (value is short shv)
        {
            cell.Value = shv;
        }
        else if (value is ushort ushv)
        {
            cell.Value = ushv;
        }
        else if (value is int iv)
        {
            cell.Value = iv;
        }
        else if (value is uint uiv)
        {
            cell.Value = uiv;
        }
        else if (value is long lv)
        {
            cell.Value = lv;
        }
        else if (value is ulong ulv)
        {
            cell.Value = ulv;
        }
        else if (value is float fv)
        {
            cell.Value = fv;
        }
        else if (value is double dv)
        {
            cell.Value = dv;
        }
        else if (value is decimal mv)
        {
            cell.Value = mv;
        }
        else
        {
            cell.Value = value?.ToString();
        }
    }
}
