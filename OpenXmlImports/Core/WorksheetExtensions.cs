using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Core
{
    public static class WorksheetExtensions
    {
        public static string GetCellText(this Cell cell, SharedStringTable sharedStrings)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                // TODO [ccb] Shared strings may also be "runs" to allow for 
                // formatting to be applied within a cell. Need to figure out how to support that.
                int index = int.Parse(cell.CellValue.Text);
                var sharedStringItem = (SharedStringItem)sharedStrings.ElementAt(index);
                return sharedStringItem.Text.Text;
            }

            return cell.CellValue == null ? string.Empty : cell.CellValue.Text;
        }

        static readonly char[] numbers = Enumerable.Range('0', 10).Select(i => (char)i).ToArray();
        public static ColumnReference Column(this StringValue stringValue)
        {
            return stringValue.Value.TrimEnd(numbers);
        }
    }
}
