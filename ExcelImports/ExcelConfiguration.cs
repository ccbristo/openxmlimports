namespace ExcelImports
{
    public static class ExcelConfiguration
    {
        public static WorkbookBuilder<T> Workbook<T>()
        {
            return new WorkbookBuilder<T>();
        }
    }
}
