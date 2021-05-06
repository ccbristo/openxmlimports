using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Types;

namespace OpenXmlImports.Core
{
    public static class WorksheetExtensions
    {
        static readonly StringType StringType = new StringType();

        public static string GetCellText(this Cell cell, SharedStringTable sharedStrings)
        {
            if (cell.DataType == null)
                return string.Empty;
            
            return (string)StringType.NullSafeGet(cell.CellValue, cell.DataType, sharedStrings);
        }

        static readonly char[] numbers = Enumerable.Range('0', 10).Select(i => (char)i).ToArray();
        public static ColumnReference Column(this StringValue stringValue)
        {
            return stringValue.Value.TrimEnd(numbers);
        }
    }
}
