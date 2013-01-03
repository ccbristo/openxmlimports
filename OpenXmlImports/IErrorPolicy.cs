using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;
using OpenXmlImports.Core;

namespace OpenXmlImports
{
    public interface IErrorPolicy
    {
        void OnInvalidFile(IEnumerable<ValidationErrorInfo> validationErrors);
        void OnMissingWorksheet(string sheetName);
        void OnMissingColumn(string columnName);
        void OnDuplicatedColumn(string columnName);
        void OnRequiredColumnViolation(string worksheetName, string columnName, ColumnReference colRef, int rowIndex);
        void OnImportComplete();
        void OnMaxLengthExceeded(ColumnReference colRef, int rowIndex, int maxLength, string columnName);
    }
}
