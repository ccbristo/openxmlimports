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
        public void Valid_Import()
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
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelImports.Tests.TestFiles.Valid_Import.xlsx"))
                result = (SingleTableHierarchy)config.Import(input);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.SingleTableItems, "result.SingleTableItems == null");
            Assert.AreEqual(3, result.SingleTableItems.Count, "SingleTableItems.Count");

            var firstItem = result.SingleTableItems[0];
            Assert.AreEqual(0, firstItem.I, "firstItem.I");
            Assert.AreEqual(christmas2012, firstItem.ADate, "firstItem.ADate");
            Assert.AreEqual("String 1", firstItem.StringField, "firstItem.StringField");

            var secondItem = result.SingleTableItems[1];
            Assert.AreEqual(654, secondItem.I, "secondItem.I");
            Assert.AreEqual(newYears2012, secondItem.ADate, "firstItem.ADate");
            Assert.IsNull(secondItem.StringField, "secondItem.StringField");

            var thirdItem = result.SingleTableItems[2];
            Assert.AreEqual(655, thirdItem.I, "thirdItem.I");
            Assert.IsNull(thirdItem.ADate, "thirdItem.I");
            Assert.AreEqual("Another string", thirdItem.StringField, "thirdItem.StringField");
        }

        [TestMethod]
        public void Missing_Worksheet()
        {
            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems");
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", "I");
            var dateColumn = singleTableItemSheet.AddColumn("A Date", "ADate");
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", "StringField");
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelImports.Tests.TestFiles.Missing_Worksheet.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have thrown an exception due to missing worksheet.");
            }
            catch (MissingWorksheetException)
            {
                // expected
            }
        }

        [TestMethod]
        public void Missing_Column()
        {
            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems");
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", "I");
            var dateColumn = singleTableItemSheet.AddColumn("A Date", "ADate");
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", "StringField");
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelImports.Tests.TestFiles.Missing_Column.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have thrown an exception due to missing column.");
            }
            catch (MissingColumnException)
            {
                // expected
            }
        }

        [TestMethod]
        public void Duplicated_Sheets()
        {
            // excel prevents this
        }

        [TestMethod]
        public void Duplicated_Columns()
        {
            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems");
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", "I");
            var dateColumn = singleTableItemSheet.AddColumn("A Date", "ADate");
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", "StringField");
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelImports.Tests.TestFiles.Duplicated_Columns.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have thrown an exception due to duplicated column.");
            }
            catch (DuplicatedColumnException)
            {
                // expected
            }
        }
    }
}
