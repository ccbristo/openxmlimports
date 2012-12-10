using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImports.Core
{
    public class ColumnConfiguration
    {
        internal ColumnConfiguration()
        { }

        public string Name { get; set; }
        public string MemberName { get; set; }

        public CellBinder GetValue(object item)
        {
            if (item == null)
                throw new ArgumentNullException("item");

            var value = item.GetType().GetMember(MemberName)
                .Single().GetPropertyOrFieldValue(item);

            CellValues cellType = InferCellType(value);

            if (value == null)
                value = string.Empty;
            else if (value is DateTime)
                value = ((DateTime)value).ToOADate().ToString();

            return new CellBinder(value.ToString(), cellType);
        }

        // TODO [ccb] This should be part of a stronger Type binding model
        private CellValues InferCellType(object value)
        {
            if (value is string || value is char)
                return CellValues.String;

            if (value is bool)
                return CellValues.Boolean;

            if (value is byte ||
                value is sbyte ||
                value is short ||
                value is ushort ||
                value is int ||
                value is uint ||
                value is long ||
                value is ulong ||
                value is float ||
                value is double ||
                value is decimal ||
                value is DateTime) // date times are stored as numbers with special styling
                return CellValues.Number;

            throw new ArgumentOutOfRangeException("value", string.Format("Could not determine cell type for {0}",
                value.GetType().Name));
        }
    }
}
