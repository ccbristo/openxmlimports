using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public interface IType
    {
        string FriendlyName { get; }
        CellValues DataType { get; }
        object NullSafeGet(CellValue cellValue, CellValues? cellType, SharedStringTable sharedStrings);
        void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings);
    }
}