using System;
using System.Linq;
using System.Reflection;

namespace ExcelImports
{
    internal static class ReflectionUtils
    {
        public static Type GetClosingInterface(this Type type, Type openInterface)
        {
            var closers = type.GetInterfaces()
                .Where(iFace => iFace.IsGenericType && iFace.GetGenericTypeDefinition() == openInterface)
                .ToList();

            if (closers.Count == 0)
                throw new ArgumentException(string.Format("{0} does not close {1}.",
                    type.Name, openInterface.Name));

            return closers[0];
        }

        public static bool ClosesInterface(this Type type, Type openInterface)
        {
            bool result = type.GetInterfaces().Any(iFace => iFace.IsGenericType && iFace.GetGenericTypeDefinition() == openInterface);
            return result;
        }

        public static bool IsPropertyOrField(this MemberInfo memberInfo)
        {
            return memberInfo.MemberType == MemberTypes.Field ||
                memberInfo.MemberType == MemberTypes.Property;
        }

        public static Type GetPropertyOrFieldType(this MemberInfo memberInfo)
        {
            FieldInfo fieldInfo;
            PropertyInfo propertyInfo;

            if ((fieldInfo = memberInfo as FieldInfo) != null)
                return fieldInfo.FieldType;

            if ((propertyInfo = memberInfo as PropertyInfo) != null)
                return propertyInfo.PropertyType;

            throw new InvalidOperationException(string.Format("{0}.{1} is not a property or field.",
                memberInfo.DeclaringType.Name, memberInfo.Name));
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
    }
}
