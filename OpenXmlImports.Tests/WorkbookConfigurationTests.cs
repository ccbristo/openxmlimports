using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlImports.Core;

namespace OpenXmlImports.Tests
{


    [TestClass]
    public class WorkbookConfigurationTests
    {
        [TestMethod]
        public void Can_Get_Added_Worksheets()
        {
            var stylesheetProvider = new DefaultStylesheetProvider();
            var config = new WorkbookConfiguration<object>(stylesheetProvider);
            var expected = new WorksheetConfiguration(typeof(object), "Sheet 1", "Sheet1", stylesheetProvider);

            config.AddWorksheet(expected);

            var actual = config.GetWorksheet(expected.SheetName);

            Assert.AreSame(expected, actual, "Different worksheet instance returned.");
        }
    }
}
