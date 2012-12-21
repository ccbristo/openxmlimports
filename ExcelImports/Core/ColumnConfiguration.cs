using System;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImports.Core
{
    public class ColumnConfiguration
    {
        internal ColumnConfiguration()
        {
            //OptionComparer = Comparer.DefaultInvariant;
            AllowNull = true;
        }

        public string Name { get; set; }
        public MemberInfo Member { get; set; }
        public NumberingFormat CellFormat { get; set; }
        public bool AllowNull { get; set; }

        // TODO [ccb] Implement these features.
        //public IEnumerable ValidValues { get; set; }
        //public bool ListValidValuesOnError { get; set; }
        //public IComparer OptionComparer { get; set; }

        internal void SetValue(object item, string text)
        {
            object value = Conversion.ExcelConvert(text, Member.GetPropertyOrFieldType());
            Member.SetPropertyOrFieldValue(item, value);
        }

        public CellBinder GetValue(object item)
        {
            if (item == null)
                throw new ArgumentNullException("item");

            var value = Member.GetPropertyOrFieldValue(item);

            CellValues cellType = InferCellType(Member);

            if (value == null)
                value = string.Empty;
            else if (value is DateTime)
                value = ((DateTime)value).ToOADate().ToString();

            return new CellBinder(value.ToString(), cellType);
        }

        // TODO [ccb] This should be part of a stronger Type binding model
        private CellValues InferCellType(MemberInfo member)
        {
            var memberType = member.GetPropertyOrFieldType();

            if (memberType.IsNullable())
                memberType = Nullable.GetUnderlyingType(memberType);

            if (memberType.In(typeof(string), typeof(char)))
                return CellValues.String;

            if (memberType == typeof(bool))
                return CellValues.Boolean;

            if (memberType.In(typeof(byte), typeof(sbyte),
                typeof(short), typeof(ushort),
                typeof(int), typeof(uint),
                typeof(long), typeof(ulong),
                typeof(float), typeof(double), typeof(decimal),
                typeof(DateTime))) // date times are stored as numbers with special styling
                return CellValues.Number;

            throw new ArgumentOutOfRangeException("member", string.Format("Could not determine cell type for {0}",
                member.GetType().Name));
        }
    }
}
