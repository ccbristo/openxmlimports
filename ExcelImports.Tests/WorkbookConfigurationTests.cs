using ExcelImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
{
    [TestClass]
    public class WorkbookConfigurationTests
    {
        [TestMethod]
        public void Can_Get_Added_Worksheets()
        {
            var config = new WorkbookConfiguration();

            var expected = new WorksheetConfiguration("Sheet 1", "Sheet1");

            config.AddWorksheet(expected);

            var actual = config.GetWorksheet(expected.SheetName);

            Assert.AreSame(expected, actual, "Different worksheet instance returned.");
        }
    }
}
