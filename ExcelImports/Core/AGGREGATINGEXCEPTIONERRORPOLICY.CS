using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;

namespace ExcelImports.Core
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
    }
}
