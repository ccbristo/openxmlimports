namespace ExcelImports.Core
{
    public class WorksheetConfiguration
    {
        public bool ExportOnly { get; set; }
        public string Name { get; set; }
    }

    public class WorksheetConfiguration<TWorksheet>
        : WorksheetConfiguration
    {
    }
}
