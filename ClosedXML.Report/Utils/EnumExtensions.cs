using System;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Report.Utils
{
    /// <summary>
    /// uses extension methods to convert enums with hypens in their names to underscore and other variants
    /// </summary>
    public static class EnumExtensions
    {
        /// <summary>
        /// Gets the description string, if available. Otherwise returns the name of the enum field
        /// LthWrapper.POS.Dollar.GetString() yields "$", an impossible control character for enums
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetString(this Enum value)
        {
            Type type = value.GetType();
            string name = GetName(type, value);
            FieldInfo field = type.GetField(name);

            if (field != null &&
                Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) is DescriptionAttribute attr)
            {
                //return the description if we have it
                name = attr.Description;
            }

            return name;
        }


        public static string[] GetStrings(Type enumType)
        {
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            if (!enumType.IsEnum)
                throw new ArgumentException("Argument 'enumType' must be enum");

            FieldInfo[] fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);
            return (from field in fieldInfo
                    let attr = (DescriptionAttribute)Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute))
                    select attr != null ? attr.Description : field.Name)
                .ToArray();
        }

        public static Array GetValues(Type enumType)
        {
#if !WindowsCE
            return Enum.GetValues(enumType);
#else
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            if (!enumType.IsEnum)
                throw new ArgumentException("Argument 'enumType' must be enum");

            object valAux = Activator.CreateInstance(enumType);
            FieldInfo[] fieldInfoArray = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);

            Array res = Array.CreateInstance(enumType, fieldInfoArray.Length);
            for (int i = 0; i < res.Length; i++)
                res.SetValue(fieldInfoArray[i].GetValue(valAux), i);
            return res;
#endif
        }

        /// <summary>
        /// Converts a string to an enum field using the string first; if that fails, tries to find a description
        /// attribute that matches. 
        /// "$".ToEnum&lt;LthWrapper.POS>() yields POS.Dollar
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="value"></param>
        /// <returns></returns>
        public static T ToEnum<T>(this string value) //, T defaultValue)
        {
            return (T)Parse(typeof(T), value, true);
        }
        private static ulong ToUInt64(object value)
        {
            switch (Convert.GetTypeCode(value))
            {
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    return (ulong)Convert.ToInt64(value, (IFormatProvider)CultureInfo.InvariantCulture);
                case TypeCode.Byte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return Convert.ToUInt64(value, (IFormatProvider)CultureInfo.InvariantCulture);
                default:
                    throw new InvalidOperationException("Invalid operation: Unknown enum type");
            }
        }

        public static string GetName(Type enumType, object value)
        {
#if WindowsCE
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            if (!enumType.IsEnum)
                throw new ArgumentException("Argument 'enumType' must be enum");
            if (value == null)
                throw new ArgumentNullException("value");

            Type type = value.GetType();
            if (type.IsEnum || type == typeof(int) || (type == typeof(short) || type == typeof(ushort)) || (type == typeof(byte) || type == typeof(sbyte) || (type == typeof(uint) || type == typeof(long))) || type == typeof(ulong))
            {
                ulong num = ToUInt64(value);
                FieldInfo[] fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);

                for (int i = 0; i < fieldInfo.Length; i++)
                    if (ToUInt64(fieldInfo[i].GetValue(value)) == num)
                        return fieldInfo[i].Name;

                throw new ArgumentOutOfRangeException("value");
            }
            else
                throw new ArgumentException("Argument 'value' must be enum base type or enum");
#else
            return Enum.GetName(enumType, value);
#endif
        }

        public static string[] GetNames(Type enumType)
        {
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            if (!enumType.IsEnum)
                throw new ArgumentException("Argument 'enumType' must be enum");

            FieldInfo[] fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);
            return fieldInfo.Select(f => f.Name).ToArray();
        }

        public static object Parse(Type enumType, string value, bool ignoreCase)
        {
            if (value == null) throw new ArgumentNullException("value");
            value = value.Trim();
            /*if (enumType == null) throw new ArgumentNullException("enumType");
            if (!enumType.IsEnum)
                throw new ArgumentException("enumType is not an Enum.");
            
            FieldInfo[] fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);
            foreach (FieldInfo field in fieldInfo)
            {
                DescriptionAttribute attr = (DescriptionAttribute) Attribute.GetCustomAttribute(field, typeof (DescriptionAttribute));
                if (string.Compare(value, attr.Description, ignoreCase) == 0 || string.Compare(value, field.Name, ignoreCase) == 0)
                    return field.GetValue(null);
            }*/
            var results = GetEnumValues(enumType, i => string.Compare(value, i, ignoreCase) == 0);
            if (results.Length != 1)
                throw new ArgumentException("value is a name, but not one of the named constants defined for the enumeration.");
            return results[0];
        }

        public static object[] GetEnumValues(Type enumType, Predicate<string> checkPredicate)
        {
            if (enumType == null) throw new ArgumentNullException("enumType");
            if (checkPredicate == null) throw new ArgumentNullException("checkPredicate");
            if (!enumType.IsEnum)
                throw new ArgumentException("enumType is not an Enum.");

            FieldInfo[] fieldInfo = enumType.GetFields(BindingFlags.Static | BindingFlags.Public);
            return (from field in fieldInfo
                    let attr = (DescriptionAttribute)Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute))
                    where (attr != null && checkPredicate(attr.Description)) || checkPredicate(field.Name)
                    select field.GetValue(null)).ToArray();
        }
    }
}
