using System;

namespace ExcelImports.Core
{
    public static class Conversion
    {
        public static object ExcelConvert(string toConvert, Type targetType)
        {
            bool isNullable = false;

            if (targetType.IsNullable())
            {
                targetType = Nullable.GetUnderlyingType(targetType);
                isNullable = true;
            }

            // excel uses OA dates, so do a special conversion here.
            if (targetType == typeof(DateTime))
                return DateTime.FromOADate(double.Parse(toConvert.ToString()));

            // unfortunately, ChangeType does NOT handle enums, so manually handle here
            if (targetType.IsEnum)
            {
                if (toConvert is string && string.IsNullOrEmpty((string)toConvert) && isNullable)
                {
                    return null;
                }

                return Enum.Parse(targetType, toConvert.ToString());
            }
            else if (isNullable && toConvert == null)
            {
                return null;
            }

            return Convert.ChangeType(toConvert, targetType);
        }
    }
}
