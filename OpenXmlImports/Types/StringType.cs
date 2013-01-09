using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

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
            // TODO [ccb] This should return shared string if "set" uses shared strings
            get { return CellValues.String; }
        }

        public object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings)
        {
            if (cellType == CellValues.SharedString)
            {
                // TODO [ccb] Shared strings may also be "runs" to allow for 
                // formatting to be applied within a cell. Need to figure out how to support that.
                int index = int.Parse(cellValue.Text);
                var sharedStringItem = (SharedStringItem)sharedStrings.ElementAt(index);
                return sharedStringItem.Text.Text;
            }

            return cellValue == null ? string.Empty : cellValue.Text;
        }

        public void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings)
        {
            // TODO [ccb] Use the shared string table here.
            cellValue.Text = (string)value;
        }
    }
}
