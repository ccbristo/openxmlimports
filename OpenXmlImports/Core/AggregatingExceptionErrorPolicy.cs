using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;

namespace OpenXmlImports.Core
{
    public class AggregatingExceptionErrorPolicy : IErrorPolicy
    {
        private readonly List<string> mErrors;

        public AggregatingExceptionErrorPolicy()
        {
            this.mErrors = new List<string>();
        }

        public void OnInvalidFile(IEnumerable<ValidationErrorInfo> validationErrors)
        {
            foreach (var error in validationErrors)
            {
                mErrors.Add(error.Description);
            }
        }

        public void OnMissingWorksheet(string sheetName)
        {
            mErrors.Add($"No sheet named \"{sheetName}\" exists in the workbook.");
        }

        public void OnMissingColumn(string columnName)
        {
            mErrors.Add($"No column named \"{columnName}\" exists in the worksheet.");
        }

        public void OnDuplicatedColumn(string columnName)
        {
            mErrors.Add($"The worksheet includes multiple columns named \"{columnName}\".");
        }

        public void OnImportComplete()
        {
            if (mErrors.Count > 0)
                throw new InvalidImportFileException(mErrors);
        }

        public void OnRequiredColumnViolation(string worksheetName, string columnName, ColumnReference colRef, uint rowIndex)
        {
            mErrors.Add(
                $"Column \"{columnName}\" on sheet \"{worksheetName}\" does not allow empty values. An empty value was found in cell {colRef}{rowIndex}.");
        }

        public void OnMaxLengthExceeded(ColumnReference colRef, uint rowIndex, int maxLength, string columnName)
        {
            mErrors.Add(
                $"The value in cell {colRef}{rowIndex} exceeds the max length of {maxLength} for column \"{columnName}\".");
        }
    }
}
