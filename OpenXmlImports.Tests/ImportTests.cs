using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlImports.Core;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class ImporterTests
    {
        private static readonly MemberInfo SingleTableHierarchyItem1sMember;
        private static readonly MemberInfo SingleTableHierarchyIMember;
        private static readonly MemberInfo SingleTableHierarchyStringFieldMember;
        private static readonly MemberInfo SingleTableHierarchyADateMember;

        static ImporterTests()
        {
            Expression<Func<SingleTableHierarchy, List<SingleTableItem>>> singleTableHierarchyItems = st => st.SingleTableItems;
            SingleTableHierarchyItem1sMember = singleTableHierarchyItems.GetMemberInfo();

            Expression<Func<SingleTableItem, int>> iProperty = sti => sti.I;
            SingleTableHierarchyIMember = iProperty.GetMemberInfo();

            Expression<Func<SingleTableItem, string>> stringField = sti => sti.StringField;
            SingleTableHierarchyStringFieldMember = stringField.GetMemberInfo();

            Expression<Func<SingleTableItem, DateTime?>> aDateProperty = sti => sti.ADate;
            SingleTableHierarchyADateMember = aDateProperty.GetMemberInfo();
        }

        [TestMethod]
        public void Valid_Import()
        {
            DateTime christmas2012 = new DateTime(2012, 12, 25, 11, 59, 30);
            DateTime newYears2012 = new DateTime(2013, 1, 1);

            var stylesheetProvider = new DefaultStylesheetProvider();
            var config = new WorkbookConfiguration<SingleTableHierarchy>(stylesheetProvider);

            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", stylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";

            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            SingleTableHierarchy result;
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Valid_Import.xlsx"))
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
        public void Can_Ignore_Rows_With_No_Cell_Values()
        {
            DateTime christmas2012 = new DateTime(2012, 12, 25, 11, 59, 30);
            DateTime newYears2012 = new DateTime(2013, 1, 1);

            var config = new WorkbookConfiguration<SingleTableHierarchy>(new DefaultStylesheetProvider());
            var stylesheetProvider = new DefaultStylesheetProvider();
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", stylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";

            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            SingleTableHierarchy result;
            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Rows_With_No_Cell_Values.xlsx"))
                result = (SingleTableHierarchy)config.Import(input);

            Assert.IsNotNull(result);
            Assert.IsNotNull(result.SingleTableItems, "result.SingleTableItems == null");
            Assert.AreEqual(1, result.SingleTableItems.Count, "SingleTableItems.Count");
        }

        [TestMethod]
        public void Error_If_Null_In_A_Required_Column()
        {
            var stylesheetProvider = new DefaultStylesheetProvider();
            var config = new WorkbookConfiguration<SingleTableHierarchy>(stylesheetProvider);
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", stylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";

            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            dateColumn.Required = true;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Null_In_A_Non_Nullable_Column.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have failed due to null in a required column.");
            }
            catch (RequiredColumnViolationException ex)
            {
                Assert.AreEqual(
                    "Column \"A Date\" on sheet \"Item 1s\" does not allow empty values. An empty value was found in cell B2.",
                    ex.Message);
            }


        }

        [TestMethod]
        public void Missing_Worksheet()
        {
            var stylesheetProvider = new DefaultStylesheetProvider();
            var config = new WorkbookConfiguration<SingleTableHierarchy>(stylesheetProvider);
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", stylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Missing_Worksheet.xlsx"))
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
            var config = new WorkbookConfiguration<SingleTableHierarchy>(new DefaultStylesheetProvider());
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", config.StylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Missing_Column.xlsx"))
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
            var config = new WorkbookConfiguration<SingleTableHierarchy>(new DefaultStylesheetProvider());
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", config.StylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Duplicated_Columns.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have thrown an exception due to duplicated column.");
            }
            catch (DuplicatedColumnException)
            {
                // expected
            }
        }

        [TestMethod]
        public void Import_Max_Length_Exceeded()
        {
            var stylesheetProvider = new DefaultStylesheetProvider();
            var config = new WorkbookConfiguration<SingleTableHierarchy>(stylesheetProvider);
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", stylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";

            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            var stringFieldColumn = singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            stringFieldColumn.MaxLength = 5;
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Max_Length_Exceeded.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);
                Assert.Fail("Import should have failed due to max length violation.");
            }
            catch (MaxLengthViolationException ex)
            {
                Assert.AreEqual(
                    "The value in cell C2 exceeds the max length of 5 for column \"String Field\".",
                    ex.Message);
            }


        }

        [TestMethod]
        public void Importer_Completes_Error_Policy()
        {
            var config = new WorkbookConfiguration<SingleTableHierarchy>(new DefaultStylesheetProvider());
            config.ErrorPolicy = new AggregatingExceptionErrorPolicy();
            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", config.StylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";
            var iColumn = singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            iColumn.Required = false;

            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            dateColumn.Required = false;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            try
            {
                SingleTableHierarchy result;
                using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Multiple_Errors.xlsx"))
                    result = (SingleTableHierarchy)config.Import(input);

                Assert.Fail("Should have thrown an exception due to errors.");
            }
            catch (InvalidImportFileException ex)
            {
                Assert.AreEqual(2, ex.Errors.Count());
            }
        }

        [TestMethod]
        public void Import_Root_Properties_And_Fields()
        {
            var config = new WorkbookConfiguration<RootPropertiesHierarchy>(new DefaultStylesheetProvider());

            var detailsSheet = new WorksheetConfiguration(typeof(RootPropertiesHierarchy), "Details", null, config.StylesheetProvider);
            detailsSheet.AddColumn("I", RootPropertiesHierarchyMembers.I);
            detailsSheet.AddColumn("Data", RootPropertiesHierarchyMembers.Data);
            config.AddWorksheet(detailsSheet);

            var singleItemSheet = new WorksheetConfiguration(typeof(Item), "Item",
                RootPropertiesHierarchyMembers.SingleItem.Name, config.StylesheetProvider);
            singleItemSheet.AddColumn("J", RootPropertiesHierarchyMembers.SingleItemMembers.J);
            singleItemSheet.AddColumn("S", RootPropertiesHierarchyMembers.SingleItemMembers.S);
            config.AddWorksheet(singleItemSheet);

            RootPropertiesHierarchy h;


            using (var input = Assembly.GetExecutingAssembly().GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Import_Root_Properties_And_Fields.xlsx"))
                h = (RootPropertiesHierarchy)config.Import(input);

            Assert.AreEqual(1, h.I, "I");
            Assert.AreEqual("Some data", h.Data, "Data");

            Assert.AreEqual(2, h.SingleItem.J, "h.SingleItem.J");
            Assert.AreEqual("S value", h.SingleItem.S, "h.SingleItem.S");
        }

        private static MemberInfo GetMember<T, TResult>(Expression<Func<T, TResult>> exp)
        {
            return exp.GetMemberInfo();
        }
    }
}
