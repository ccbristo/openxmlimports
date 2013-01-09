using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class EnumStringType : IType
    {
        public Type EnumType { get; private set; }
        private readonly StringType StringType = new StringType();

        public EnumStringType(Type enumType)
        {
            if (!enumType.Is<Enum>())
                throw new ArgumentException("enumType", "enumType must be an enumeration.");

            this.EnumType = enumType.IsNullable() ? Nullable.GetUnderlyingType(enumType) : enumType;
            this.FriendlyName = new CamelCaseNamingConvention().GetName(EnumType);
        }

        public string FriendlyName { get; private set; }

        public CellValues DataType
        {
            get { return StringType.DataType; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            string s = (string)StringType.NullSafeGet(cellValue, cellType, sharedStrings);

            if (string.IsNullOrWhiteSpace(s))
                return null;

            return Enum.Parse(EnumType, s);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            value = value != null ? value.ToString() : null;
            StringType.NullSafeSet(cellValue, value, sharedStrings);
        }
    }

    public class EnumStringType<T> : EnumStringType
    {
        public EnumStringType()
            : base(typeof(T))
        { }
    }
}
