using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Report.Utils
{
    /// <summary>
    /// uses extension methods to convert enums with hyphens in their names to underscore and other variants
    /// </summary>
    public static class EnumExtensions
    {
        public static object Parse(Type enumType, string value, bool ignoreCase)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            value = value.Trim();
            var results = GetEnumValues(enumType, i => string.Compare(value, i, ignoreCase) == 0);
            if (results.Length != 1)
                throw new ArgumentException(
                    "value is a name, but not one of the named constants defined for the enumeration.");
            return results[0];
        }

        private static object[] GetEnumValues(Type enumType, Predicate<string> checkPredicate)
        {
            if (enumType == null) throw new ArgumentNullException(nameof(enumType));
            if (checkPredicate == null) throw new ArgumentNullException(nameof(checkPredicate));
            if (!enumType.IsEnum)
                throw new ArgumentException("enumType is not an Enum.");

            var fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);
            return (from field in fieldInfo
                let attr = (DescriptionAttribute)Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute))
                where (attr != null && checkPredicate(attr.Description)) || checkPredicate(field.Name)
                select field.GetValue(null)).ToArray();
        }
    }
}