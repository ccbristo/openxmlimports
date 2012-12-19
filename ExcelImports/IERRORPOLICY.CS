using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;

namespace ExcelImports
{
    public interface IErrorPolicy
    {
        void OnInvalidFile(IEnumerable<ValidationErrorInfo> validationErrors);
        void OnMissingWorksheet(string sheetName);
        void OnMissingColumn(string columnName);
        void OnDuplicatedColumn(string columnName);
        void OnImportComplete();
    }
}
