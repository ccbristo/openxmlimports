using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class ShortType : IType
    {
        public string FriendlyName
        {
            get { return "whole number"; }
        }

        public CellValues DataType
        {
            get { return CellValues.Number; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.Text))
                return null;

            return short.Parse(cellValue.Text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
