﻿using System.Collections.Generic;
using DocumentFormat.OpenXml.Validation;
using OpenXmlImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
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
                Assert.Fail("InvalidImportFileException was not thrown when the import was completed with errors.");
            }
            catch (InvalidImportFileException)
            {
                // expected
            }
        }
    }
}
