using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Core;

namespace OpenXmlImports.Types
{
    public class SingleType : IType
    {
        public string FriendlyName
        {
            get { return "number"; }
        }

        public CellValues DataType
        {
            get { return CellValues.Number; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.Text))
                return null;

            string text = cellValue.Text;
            if (cellType == CellValues.SharedString)
                text = sharedStrings.GetText(cellValue);

            return float.Parse(text.TrimStart('\''));
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
