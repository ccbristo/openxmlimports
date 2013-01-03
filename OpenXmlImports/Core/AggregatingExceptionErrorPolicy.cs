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
            mErrors.Add(string.Format("No sheet named \"{0}\" exists in the workbook.", sheetName));
        }

        public void OnMissingColumn(string columnName)
        {
            mErrors.Add(string.Format("No column named \"{0}\" exists in the worksheet.", columnName));
        }

        public void OnDuplicatedColumn(string columnName)
        {
            mErrors.Add(string.Format("The worksheet includes multiple columns named \"{0}\".", columnName));
        }

        public void OnImportComplete()
        {
            if (mErrors.Count > 0)
                throw new InvalidImportFileException(mErrors);
        }

        public void OnNullableColumnViolation(string worksheetName, string columnName, ColumnReference colRef, int rowIndex)
        {
            mErrors.Add(string.Format("Column \"{0}\" on sheet \"{1}\" does not allow empty values. An empty value was found in cell {2}{3}.",
                columnName, worksheetName, colRef, rowIndex));
        }

        public void OnMaxLengthExceeded(ColumnReference colRef, int rowIndex, int maxLength, string columnName)
        {
            mErrors.Add(string.Format("The value in cell {0}{1} exceeds the max length of {2} for column \"{3}\".",
                colRef, rowIndex, maxLength, columnName));
        }
    }
}
