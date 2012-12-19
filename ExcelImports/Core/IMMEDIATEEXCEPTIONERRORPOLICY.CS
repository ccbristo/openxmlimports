using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;

namespace ExcelImports.Core
{
    public class ImmediateExceptionErrorPolicy : IErrorPolicy
    {
        public void OnInvalidFile(IEnumerable<ValidationErrorInfo> validationErrors)
        {
            var errors = validationErrors.Aggregate(new StringBuilder(),
                (sb, error) => sb.AppendLine(error.Description),
                sb => sb.ToString());

            throw new InvalidMCContentException(errors);
        }

        public void OnMissingWorksheet(string sheetName)
        {
            throw new MissingWorksheetException("No sheet named \"{0}\" was found in the workbook.",
                sheetName);
        }

        public void OnMissingColumn(string columnName)
        {
            throw new MissingColumnException("No column named \"{0}\" could be found.",
                        columnName);
        }

        public void OnDuplicatedColumn(string columnName)
        {
            throw new DuplicatedColumnException("Found multiple columns named \"{0}\".", columnName);
        }

        public void OnImportComplete()
        {
            // nop
        }
    }
}
