using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Core;

namespace OpenXmlImports.Types
{
    public class StringType : IType
    {
        public string FriendlyName
        {
            get { return "text block"; }
        }

        public CellValues DataType
        {
            get { return CellValues.SharedString; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellType == CellValues.SharedString)
                return sharedStrings.GetText(cellValue);

            return cellValue == null ? string.Empty : cellValue.Text;
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            int? index = GetIndex((string)value, sharedStrings);

            if (index != null)
            {
                cellValue.Text = index.Value.ToString();
                return;
            }

            var item = new SharedStringItem(new Text((string)value));
            sharedStrings.AppendChild(item);
            cellValue.Text = (sharedStrings.Count() - 1).ToString();
        }

        static int? GetIndex(string text, SharedStringTable sharedStrings)
        {
            int index = -1;
            bool found = false;

            sharedStrings.Elements<SharedStringItem>()
                .SkipWhile(item =>
                {
                    index++;
                    found = StringComparer.Ordinal.Equals(item.InnerText, text);
                    return !found;
                });

            return found ? (int?)index : null;
        }
    }
}
