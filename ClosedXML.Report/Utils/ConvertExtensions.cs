using System;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Report.Utils
{
    public static class ConvertExtensions
    {
        public static T As<T>(this object obj)
        {
            return (T)obj;
        }

        public static int AsInt(this string value)
        {
            return AsInt(value, 0);
        }

        public static int AsInt(this string value, int def, CultureInfo culture = null)
        {
            culture = culture ?? CultureInfo.CurrentCulture;
            if (!string.IsNullOrEmpty(value) &&
                int.TryParse(value, NumberStyles.AllowThousands, culture, out var result))
            {
                return result;
            }

            return value.ExtractNumber(def);
        }

        public static int AsInt(this bool value)
        {
            return value ? 1 : 0;
        }

        public static bool AsBool(this object value)
        {
            if (value == null)
                return false;

            if (value is bool)
                return (bool)value;

            return value.ToString().AsBool();
        }

        public static bool AsBool(this string value)
        {
            if (string.IsNullOrEmpty(value))
                return false;
            switch (value.ToLower())
            {
                case "0":
                case "0.0":
                case "0,0":
                case "-":
                case "false":
                case "ложь":
                case "нет":
                case "no":
                case "not":
                    return false;
                default:
                    return true;
            }
        }

        public static T AsEnum<T>(this string value, T def)
        {
            try
            {
                if (string.IsNullOrEmpty(value))
                    return def;

                return (T)EnumExtensions.Parse(typeof(T), value, true);
            }
            catch (ArgumentException)
            {
                return def;
            }
        }

        public static DateTime AsDateTime(this string value)
        {
            return AsDateTime(value, CultureInfo.CurrentCulture, DateTime.MinValue);
        }

        public static DateTime AsDateTime(this string value, CultureInfo culture, DateTime def)
        {
            string[] fmts = { "yyyy-MM-dd HH:mm:ssZ" };
            fmts = fmts.Union(culture.DateTimeFormat.GetAllDateTimePatterns()).ToArray();
            if (Equals(culture, CultureInfo.InvariantCulture))
                fmts = fmts.Union(CultureInfo.CurrentCulture.DateTimeFormat.GetAllDateTimePatterns()).ToArray();
            else
                fmts = fmts.Union(CultureInfo.InvariantCulture.DateTimeFormat.GetAllDateTimePatterns()).ToArray();
            if (!string.IsNullOrEmpty(value) &&
                DateTime.TryParseExact(value, fmts, culture, DateTimeStyles.None, out var result))
            {
                return result;
            }

            return def;
        }

        public static decimal AsDecimal(this string value)
        {
            return AsDecimal(value, 0);
        }

        public static decimal AsDecimal(this string value, decimal def, CultureInfo culture = null)
        {
            culture = culture ?? CultureInfo.CurrentCulture;
            value = ReplaceNumberFormat(value);
            try
            {
                if (!string.IsNullOrEmpty(value))
                    return Convert.ToDecimal(value, culture);
            }
            catch (FormatException)
            {
            }
            return def;
        }

        public static double AsDouble(this string value)
        {
            return AsDouble(value, 0);
        }

        public static double AsDouble(this string value, double def, CultureInfo culture = null)
        {
            value = ReplaceNumberFormat(value);
            if (!string.IsNullOrEmpty(value) &&
                double.TryParse(value, NumberStyles.AllowThousands, culture, out var result))
            {
                return result;
            }

            return def;
        }

        public static float AsFloat(this string value)
        {
            return AsFloat(value, 0);
        }

        public static float AsFloat(this string value, float def, CultureInfo culture = null)
        {
            value = ReplaceNumberFormat(value);
            if (!string.IsNullOrEmpty(value) &&
                float.TryParse(value, NumberStyles.AllowThousands, culture, out var result))
            {
                return result;
            }

            return def;
        }

        public static string ReplaceNumberFormat(string value)
        {
            var sep = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            value = value.Replace(".", sep).Replace(",", sep);
            return value;
        }

        public static Guid AsGuid(this string value)
        {
            return AsGuid(value, Guid.Empty);
        }

        public static Guid AsGuid(this string value, Guid def)
        {
            try
            {
                if (string.IsNullOrEmpty(value))
                    return def;

                return new Guid(value);
            }
            catch
            {
                return def;
            }
        }

        public static T ChangeType<T>(object value)
        {
            var t = typeof(T);

            if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                return value != null
                    ? (T)Convert.ChangeType(value, Nullable.GetUnderlyingType(t), CultureInfo.CurrentCulture)
                    : default(T);
            else
                return (T)Convert.ChangeType(value, t, CultureInfo.CurrentCulture);
        }

        public static object ChangeType(this object value, Type conversion)
        {
            return value.ChangeType(conversion, CultureInfo.CurrentCulture);
        }

        public static object ChangeType(this object value, Type conversion, CultureInfo culture)
        {
            var t = conversion;

            if (value is string stringValue)
            {
                if (typeof(int) == conversion)
                    return stringValue.AsInt(0, culture);
                else if (typeof(double) == conversion)
                    return stringValue.AsDouble(0, culture);
                else if (typeof(float) == conversion)
                    return stringValue.AsFloat(0, culture);
                else if (typeof(decimal) == conversion)
                    return stringValue.AsDecimal(0, culture);
            }

            if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                return value != null
                    ? Convert.ChangeType(value, Nullable.GetUnderlyingType(t), culture)
                    : t.GetDefault();
            }
            else
            {
                return Convert.ChangeType(value, t, culture);
            }
        }
    }
}
