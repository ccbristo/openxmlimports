using System;

namespace ExcelImports.Core
{
    public static class Conversion
    {
        public static object ExcelConvert(string toConvert, Type targetType)
        {
            Type trueTargetType = null;
            bool isNullable = false;

            if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == (typeof(Nullable<>)))
            {
                trueTargetType = Nullable.GetUnderlyingType(targetType);
                isNullable = true;
            }
            else
            {
                trueTargetType = targetType;
            }

            // excel uses OA dates, so do a special conversion here.
            if (trueTargetType == typeof(DateTime))
                return DateTime.FromOADate(double.Parse(toConvert.ToString()));

            // unfortunately, ChangeType does NOT handle enums, so manually handle here
            if (trueTargetType.IsEnum)
            {
                if (toConvert is string && string.IsNullOrEmpty((string)toConvert) && isNullable)
                {
                    return null;
                }

                return Enum.Parse(trueTargetType, toConvert.ToString());
            }
            else if (isNullable && toConvert == null)
            {
                return null;
            }

            return Convert.ChangeType(toConvert, trueTargetType);
        }
    }
}
