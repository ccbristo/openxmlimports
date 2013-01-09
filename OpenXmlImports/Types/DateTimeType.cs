using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public class DateTimeType : IType
    {
        public string FriendlyName
        {
            get { return "date"; }
        }

        public CellValues DataType
        {
            get { return CellValues.Number; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellValue == null || string.IsNullOrEmpty(cellValue.Text))
                return null;

            return DateTime.FromOADate(double.Parse(cellValue.Text));
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
