using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Report.Utils
{
    public static class TypeExtensions
    {
        private static readonly HashSet<Type> NumericTypes = new()
        {
            typeof(byte), typeof(sbyte), typeof(short), typeof(ushort), typeof(int),
            typeof(uint), typeof(long), typeof(ulong), typeof(double), typeof(decimal), typeof(float)
        };

        public static bool IsPrimitive(this Type type)
        {
            if (type == typeof(string) || type.IsEnum || type == typeof(DateTime)
                || type == typeof(Guid) || type == typeof(decimal) || type.IsNullable()) return true;
            return (type.IsValueType & type.IsPrimitive);
        }

        public static Type GetItemType(this Type type)
        {
            if (!typeof(IEnumerable).IsAssignableFrom(type))
                throw new ArgumentException("Argument 'type' should be inherited from IEnumerable.");

            if (type.IsArray)
                return type.GetElementType();
            else
            {
                var genericType = (from intType in type.GetInterfaces()
                    where intType.IsGenericType && intType.GetGenericTypeDefinition() == typeof(IEnumerable<>)
                    select intType.GetGenericArguments()[0]).FirstOrDefault();

                if (genericType != null)
                    return genericType;
                else
                    return typeof(object);
            }
        }

        public static Type GetItemType(this IEnumerable sourceList)
        {
            var type = sourceList.GetType();

            var result = type.GetItemType();
            if (result != typeof(object))
                return result;

            var item = sourceList.Cast<object>().FirstOrDefault(x => x != null);
            if (item == null)
                return typeof(object);
            var itemType = item.GetType();
            return !sourceList.IsAllItemsAssignableFrom(itemType)
                ? typeof(object)
                : itemType;
        }

        public static bool IsAllItemsAssignableFrom(this IEnumerable list, Type type)
        {
            return list.Cast<object>().All(item =>
                type.IsValueType
                    ? item != null && item.GetType().IsAssignableFrom(type)
                    : item == null || item.GetType().IsAssignableFrom(type));
        }

        internal static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type) ||
                   NumericTypes.Contains(Nullable.GetUnderlyingType(type));
        }

        public static bool IsNullable(this Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        public static object GetDefault(this Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }
    }
}