using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class DateTimeType : IType
    {
        public string FriendlyName => "date";

        public CellValues DataType => CellValues.Number;

        public object NullSafeGet(string text)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            return DateTime.FromOADate(double.Parse(text));
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            if (value != null)
            {
                cellValue.Text = ((DateTime)value).ToOADate().ToString();
            }
        }
    }
}
