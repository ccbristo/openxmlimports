using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Core
{
    public static class WorksheetExtensions
    {
        public static string GetCellText(this Cell cell, SharedStringTable sharedStrings)
        {
            if (cell.DataType == null)
                return string.Empty;

            if (cell.DataType?.Value == CellValues.SharedString)
                return sharedStrings.GetText(cell.CellValue);

            return cell.CellValue.Text;
        }

        private static readonly char[] Numbers = Enumerable.Range('0', 10).Select(i => (char)i).ToArray();
        public static ColumnReference Column(this StringValue stringValue)
        {
            var columnOnly = stringValue.Value.TrimEnd(Numbers);
            return new ColumnReference(columnOnly);
        }

        internal static string GetText(this SharedStringTable sharedStrings, CellValue cellValue)
        {
            // TODO [ccb] Shared strings may also be "runs" to allow for 
            // formatting to be applied within a cell. Need to figure out how to support that.
            int index = int.Parse(cellValue.Text);
            var sharedStringItem = (SharedStringItem)sharedStrings.ElementAt(index);
            return sharedStringItem.Text.Text;
        }
    }
}
