using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class GivenASheetWithBordersAroundCellsWithNoContent
    {
        [TestClass]
        public class WhenItIsImported
        {
            private readonly SingleTableHierarchy _result;

            public WhenItIsImported()
            {
                var config = OpenXmlConfiguration.Workbook<SingleTableHierarchy>().Create();
                using (var input = Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Empty_Header_Cells.xlsx"))
                {
                    _result = config.Import(input);
                }
            }

            [TestMethod]
            public void TheDataIsReadSuccessfully()
            {
                Assert.AreEqual(3, _result.SingleTableItems.Count);
            }
        }
    }
}