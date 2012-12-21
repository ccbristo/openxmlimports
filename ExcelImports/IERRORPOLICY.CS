using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;
using ExcelImports.Core;

namespace ExcelImports
{
    public interface IErrorPolicy
    {
        void OnInvalidFile(IEnumerable<ValidationErrorInfo> validationErrors);
        void OnMissingWorksheet(string sheetName);
        void OnMissingColumn(string columnName);
        void OnDuplicatedColumn(string columnName);
        void OnNullableColumnViolation(string worksheetName, string columnName, ColumnReference colRef, int rowIndex);
        void OnImportComplete();
    }
}
