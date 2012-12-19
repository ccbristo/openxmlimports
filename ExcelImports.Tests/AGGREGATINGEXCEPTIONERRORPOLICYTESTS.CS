using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;
using ExcelImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
{
    [TestClass]
    public class AggregatingExceptionErrorPolicyTests
    {
        [TestMethod]
        public void AggregatingExceptionErrorPolicy_No_Errors()
        {
            var policy = new AggregatingExceptionErrorPolicy();
            policy.OnImportComplete();
        }

        [TestMethod]
        public void AggregatingExceptionErrorPolicy_Aggregates_Errors()
        {
            var policy = new AggregatingExceptionErrorPolicy();

            policy.OnDuplicatedColumn("column 1");
            policy.OnMissingColumn("column 2");
            policy.OnMissingWorksheet("sheet 1");
            policy.OnInvalidFile(new List<ValidationErrorInfo>()
                {
                    new ValidationErrorInfo()
                });

            try
            {
                policy.OnImportComplete();
            }
            catch (InvalidImportFileException)
            {
                // expected
            }
        }
    }
}
