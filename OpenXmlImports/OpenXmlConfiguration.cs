namespace OpenXmlImports
{
    public static class OpenXmlConfiguration
    {
        public static WorkbookBuilder<T, DefaultStylesheetProvider> Workbook<T>()
        {
            return new WorkbookBuilder<T, DefaultStylesheetProvider>(new DefaultStylesheetProvider());
        }

        public static WorkbookBuilder<T, TStylesheet> Workbook<T, TStylesheet>(TStylesheet stylesheet)
            where TStylesheet : IStylesheetProvider
        {
            return new WorkbookBuilder<T, TStylesheet>(stylesheet);
        }
    }
}
