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
        void OnRequiredColumnViolation(string worksheetName, string columnName, ColumnReference colRef, uint rowIndex);
        void OnImportComplete();
        void OnMaxLengthExceeded(ColumnReference colRef, uint rowIndex, int maxLength, string columnName);
    }
}
