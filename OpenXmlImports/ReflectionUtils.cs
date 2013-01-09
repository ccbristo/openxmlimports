using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OpenXmlImports
{
    internal static class ReflectionUtils
    {
        public static Type GetClosingInterface(this Type type, Type openInterface)
        {
            var closers = type.GetAllInterfaces()
                .Where(iFace => iFace.IsGenericType && iFace.GetGenericTypeDefinition() == openInterface)
                .ToList();

            if (closers.Count == 0)
                throw new ArgumentException(string.Format("{0} does not close {1}.",
                    type.Name, openInterface.Name));

            return closers[0];
        }

        public static bool ClosesInterface(this Type type, Type openInterface)
        {
            bool result = type.GetAllInterfaces().Any(iFace => iFace.IsGenericType && iFace.GetGenericTypeDefinition() == openInterface);
            return result;
        }

        public static IEnumerable<Type> GetAllInterfaces(this Type type)
        {
            var interfaces = new List<Type>();

            if (type.IsInterface)
                interfaces.Add(type);

            interfaces.AddRange(type.GetInterfaces());

            return interfaces;
        }

        public static bool IsPropertyOrField(this MemberInfo memberInfo)
        {
            return memberInfo.MemberType == MemberTypes.Field ||
                memberInfo.MemberType == MemberTypes.Property;
        }

        public static Type GetMemberType(this MemberInfo memberInfo)
        {
            FieldInfo fieldInfo;
            if ((fieldInfo = memberInfo as FieldInfo) != null)
                return fieldInfo.FieldType;

            PropertyInfo propertyInfo;
            if ((propertyInfo = memberInfo as PropertyInfo) != null)
                return propertyInfo.PropertyType;

            Type type;
            if ((type = memberInfo as Type) != null)
                return type;

            throw new InvalidOperationException(string.Format("{0}.{1} is not a property or field.",
                memberInfo.DeclaringType.Name, memberInfo.Name));
        }

        public static bool Is<T>(this Type type)
        {
            Type targetType = type.IsNullable() ? Nullable.GetUnderlyingType(type) : type;
            return typeof(T).IsAssignableFrom(targetType);
        }

        public static object GetPropertyOrFieldValue(this MemberInfo memberInfo, object source)
        {
            PropertyInfo propertyInfo;
            FieldInfo fieldInfo;

            if ((propertyInfo = memberInfo as PropertyInfo) != null)
                return propertyInfo.GetValue(source, null);

            if ((fieldInfo = memberInfo as FieldInfo) != null)
                return fieldInfo.GetValue(source);

            throw new InvalidOperationException(string.Format("{0}.{1} is not a property or field.",
                memberInfo.DeclaringType.Name, memberInfo.Name));
        }

        public static void SetPropertyOrFieldValue(this MemberInfo member, object instance, object value)
        {
            FieldInfo fieldInfo;
            PropertyInfo propertyInfo;

            if ((fieldInfo = member as FieldInfo) != null)
                fieldInfo.SetValue(instance, value);
            else if ((propertyInfo = member as PropertyInfo) != null)
                propertyInfo.SetValue(instance, value, null);
            else
                throw new InvalidOperationException(string.Format("{0}.{1} is not a property or field.",
                    member.DeclaringType.Name, member.Name));
        }

        public static bool IsNullable(this Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        public static bool IsTerminal(this Type type)
        {
            if (type.IsNullable())
                type = Nullable.GetUnderlyingType(type);

            return type.IsPrimitive ||
                type.Is<Enum>() ||
                type.In(typeof(string), typeof(decimal), typeof(DateTime), typeof(Guid));

        }
    }
}
