using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class DoubleType : IType
    {
        public string FriendlyName => "number";

        public CellValues DataType => CellValues.Number;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return double.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}