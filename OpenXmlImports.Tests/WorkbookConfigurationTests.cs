using OpenXmlImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class WorkbookConfigurationTests
    {
        [TestMethod]
        public void Can_Get_Added_Worksheets()
        {
            var config = new WorkbookConfiguration(typeof(object));
            var stylesheetProvider = new DefaultStylesheetProvider();
            var expected = new WorksheetConfiguration(typeof(object), "Sheet 1", "Sheet1", stylesheetProvider);

            config.AddWorksheet(expected);

            var actual = config.GetWorksheet(expected.SheetName);

            Assert.AreSame(expected, actual, "Different worksheet instance returned.");
        }
    }
}
