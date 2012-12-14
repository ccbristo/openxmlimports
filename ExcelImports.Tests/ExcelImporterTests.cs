using System;
using System.Reflection;
using ExcelImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
{
    [TestClass]
    public class ExcelImporterTests
    {
        [TestMethod]
        public void Import()
        {
            DateTime christmas2012 = new DateTime(2012, 12, 25, 11, 59, 30);
            DateTime newYears2012 = new DateTime(2013, 1, 1);

            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems");
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", "I");
            var dateColumn = singleTableItemSheet.AddColumn("A Date", "ADate");
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", "StringField");
            config.AddWorksheet(singleTableItemSheet);

            SingleTableHierarchy result;
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelImports.Tests.TestFiles.Basic_Import.xlsx"))
                result = (SingleTableHierarchy)config.Import(input);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.SingleTableItems, "result.SingleTableItems == null");
            Assert.AreEqual(2, result.SingleTableItems.Count, "SingleTableItems.Count");

            var firstItem = result.SingleTableItems[0];
            Assert.AreEqual(0, firstItem.I, "firstItem.I");
            Assert.AreEqual("String 1", firstItem.StringField, "firstItem.StringField");

            var secondItem = result.SingleTableItems[1];
            Assert.AreEqual(654, secondItem.I, "secondItem.I");
            Assert.IsNull(secondItem.StringField, "secondItem.StringField");
        }
    }
}
