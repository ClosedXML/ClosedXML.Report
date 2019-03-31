using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ClosedXML.Report.Utils
{
    public static class TypeExtensions
    {
        private static readonly HashSet<Type> NumericTypes = new HashSet<Type>
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

            var item = sourceList.Cast<object>().First();
            if (item == null)
                return typeof(object);
            return item.GetType();
        }

        internal static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type) ||
                   NumericTypes.Contains(Nullable.GetUnderlyingType(type));
        }

        public static Type ToType(this TypeCode code)
        {
            switch (code)
            {
                case TypeCode.Boolean:
                    return typeof(bool);

                case TypeCode.Byte:
                    return typeof(byte);

                case TypeCode.Char:
                    return typeof(char);

                case TypeCode.DateTime:
                    return typeof(DateTime);

                case TypeCode.DBNull:
                    return typeof(DBNull);

                case TypeCode.Decimal:
                    return typeof(decimal);

                case TypeCode.Double:
                    return typeof(double);

                case TypeCode.Empty:
                    return null;

                case TypeCode.Int16:
                    return typeof(short);

                case TypeCode.Int32:
                    return typeof(int);

                case TypeCode.Int64:
                    return typeof(long);

                case TypeCode.Object:
                    return typeof(object);

                case TypeCode.SByte:
                    return typeof(sbyte);

                case TypeCode.Single:
                    return typeof(Single);

                case TypeCode.String:
                    return typeof(string);

                case TypeCode.UInt16:
                    return typeof(UInt16);

                case TypeCode.UInt32:
                    return typeof(UInt32);

                case TypeCode.UInt64:
                    return typeof(UInt64);
            }

            return null;
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

        public static PropertyInfo GetPropertyInfo(this Type type, string fieldName, BindingFlags flags = BindingFlags.Instance | BindingFlags.Public)
        {
            var property = type.GetProperty(fieldName, flags);
#if NET35
            if (property == null) throw new MissingMemberException(string.Format("Invalid property {0}.{1}", type.Name, fieldName));
#else
            if (property == null) throw new MissingMemberException(type.Name, fieldName);
#endif
            return property;
        }

        public static bool IsAssignableToGenericType(this Type givenType, Type genericType)
        {
            var interfaceTypes = givenType.GetInterfaces();

            if (interfaceTypes.Any(it => it.IsGenericType && it.GetGenericTypeDefinition() == genericType))
                return true;

            if (givenType.IsGenericType && givenType.GetGenericTypeDefinition() == genericType)
                return true;

            Type baseType = givenType.BaseType;
            return baseType != null && IsAssignableToGenericType(baseType, genericType);
        }

#if !WindowsCE
        public static PropertyInfo GetPropertyInfo<TSource, TProperty>(
            this Type type,
            Expression<Func<TSource, TProperty>> propertyLambda)
        {
            if (!(propertyLambda.Body is MemberExpression member))
                throw new ArgumentException(string.Format(
                    "Expression '{0}' refers to a method, not a property.",
                    propertyLambda.ToString()));

            PropertyInfo propInfo = member.Member as PropertyInfo;
            if (propInfo == null)
                throw new ArgumentException(string.Format(
                    "Expression '{0}' refers to a field, not a property.",
                    propertyLambda.ToString()));

            if (type != propInfo.ReflectedType &&
                !type.IsSubclassOf(propInfo.ReflectedType))
                throw new ArgumentException(string.Format(
                    "Expresion '{0}' refers to a property that is not from type {1}.",
                    propertyLambda.ToString(),
                    type));

            return propInfo;
        }
#endif

        public static object GetInstanceField(this Type type, object instance, string fieldName)
        {
            BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
                | BindingFlags.Static;
            FieldInfo field = type.GetField(fieldName, bindFlags);
            return field.GetValue(instance);
        }

        public static object Invoke(this Type type, object instance, string methodName, params object[] methodParams)
        {
            MethodInfo dynMethod = type.GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance);
            return dynMethod.Invoke(instance, methodParams);
        }

#if NET20 || NET35 || WindowsCE 
        public static T GetCustomAttribute<T>(this MemberInfo member, bool inherit)
            where T : Attribute
        {
            return (T)Attribute.GetCustomAttribute(member, typeof (T), inherit);
        }

        public static T GetCustomAttribute<T>(this ParameterInfo member, bool inherit)
            where T : Attribute
        {
            return (T)Attribute.GetCustomAttribute(member, typeof (T), inherit);
        }
#endif
    }
}
