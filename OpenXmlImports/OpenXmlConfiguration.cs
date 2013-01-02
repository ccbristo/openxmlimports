namespace OpenXmlImports
{
    public static class OpenXmlConfiguration
    {
        public static WorkbookBuilder<T> Workbook<T>()
        {
            return new WorkbookBuilder<T>();
        }
    }
}
