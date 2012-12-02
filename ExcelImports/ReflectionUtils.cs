using System;
using System.Linq;
using System.Reflection;

namespace ExcelImports
{
    public static class ReflectionUtils
    {
        public static bool ClosesInterface(this MemberInfo memberInfo, Type openInterface)
        {
            var type = memberInfo.GetPropertyOrFieldType();
            bool result = type.GetInterfaces().Any(iFace => iFace.IsGenericType && iFace.GetGenericTypeDefinition() == openInterface);
            return result;
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
    }
}
