using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class BooleanType : IType
    {
        public string FriendlyName => "boolean";

        public CellValues DataType => CellValues.String;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return bool.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
