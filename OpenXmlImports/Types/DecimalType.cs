using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Core;

namespace OpenXmlImports.Types
{
    public class DecimalType : IType
    {
        public string FriendlyName => "number";

        public CellValues DataType => CellValues.Number;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return decimal.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
