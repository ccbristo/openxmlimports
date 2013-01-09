using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class BooleanType : IType
    {
        public string FriendlyName
        {
            get { return "boolean"; }
        }

        public CellValues DataType
        {
            get { return CellValues.String; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.Text))
                return null;

            return bool.Parse(cellValue.Text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
