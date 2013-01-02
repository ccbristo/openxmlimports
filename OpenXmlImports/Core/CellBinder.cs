using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Core
{
    public class CellBinder
    {
        public string Value { get; private set; }
        public CellValues CellType { get; private set; }

        public CellBinder(string value, CellValues cellType)
        {
            this.Value = value;
            this.CellType = cellType;
        }
    }
}
