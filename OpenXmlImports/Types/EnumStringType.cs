using System;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Core;

namespace OpenXmlImports.Types
{
    public class EnumStringType : IType
    {
        private readonly Type EnumType;
        private readonly StringType StringType = new StringType();

        private static readonly MethodInfo EnumTryParseGenericMethodInfo;

        static EnumStringType()
        {
            EnumTryParseGenericMethodInfo = typeof(Enum).GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Single(m => m.IsGenericMethod && m.Name == "TryParse" && m.GetParameters().Length == 3);
        }

        public EnumStringType(Type enumType)
        {
            if (!enumType.Is<Enum>())
                throw new ArgumentException("enumType", "enumType must be an enumeration.");

            this.EnumType = enumType.IsNullable() ? Nullable.GetUnderlyingType(enumType) : enumType;
            this.FriendlyName = new CamelCaseNamingConvention().GetName(EnumType);
        }

        public string FriendlyName { get; }

        public CellValues DataType => StringType.DataType;

        public object NullSafeGet(string text)
        {
            return Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            value = Format(value);
            StringType.NullSafeSet(cellValue, value, sharedStrings);
        }

        internal object Format(object value)
        {
            return value == null ? null : CamelCaseFormatter.Format(value.ToString());
        }

        internal object Parse(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return null;

            if (DynamicTryParse(s, out Enum e))
                return e;

            var formattedTextValues = Enum.GetValues(EnumType)
                .Cast<Enum>()
                .ToDictionary(val => CamelCaseFormatter.Format(val.ToString()), val => val);

            if (formattedTextValues.TryGetValue(s, out e))
                return e;

            return new FormatException($"Could not convert {s} to an {EnumType.Name}.");
        }

        private bool DynamicTryParse(string s, out Enum @enum)
        {
            @enum = default(Enum);

            var closedMethod = EnumTryParseGenericMethodInfo.MakeGenericMethod(EnumType);
            var parameters = new object[] { s, true, null };
            var result = (bool)closedMethod.Invoke(null, parameters);

            if (result)
                @enum = (Enum)parameters[2]; // get the out parameter

            return result;

        }
    }
}