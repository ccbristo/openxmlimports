using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Core
{
    public interface IStylesheetProvider
    {
        Stylesheet Stylesheet { get; }

        NumberingFormat DateFormat { get; }
        NumberingFormat DateTimeFormat { get; }
        NumberingFormat TimeFormat { get; }
        uint GetStyleIndex(NumberingFormat format);
    }
}
