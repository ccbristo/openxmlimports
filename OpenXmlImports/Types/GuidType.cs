using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class GuidType : IType
    {
        public string FriendlyName => "guid";

        public CellValues DataType => CellValues.String;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return Guid.Parse(text);
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
                cellValue.Text = value.ToString();
        }
    }
}
