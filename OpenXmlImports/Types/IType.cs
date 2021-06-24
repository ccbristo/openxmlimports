using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Types
{
    public interface IType
    {
        string FriendlyName { get; }
        CellValues DataType { get; }
        object NullSafeGet(string text);
        void NullSafeSet(CellValue cellValue, object value, SharedStringTable sharedStrings);
    }
}