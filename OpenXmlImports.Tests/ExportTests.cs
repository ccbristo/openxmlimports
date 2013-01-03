using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlImports.Core;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class ExporterTests
    {
        private static readonly MemberInfo SingleTableHierarchyItem1sMember;
        private static readonly MemberInfo SingleTableHierarchyIMember;
        private static readonly MemberInfo SingleTableHierarchyStringFieldMember;
        private static readonly MemberInfo SingleTableHierarchyADateMember;

        static ExporterTests()
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
        public void Can_Export_Worksheets()
        {
            var config = new WorkbookConfiguration(typeof(SimpleHierarchy));
            var dataSource = new SimpleHierarchy()
            {
                I = 1,
                S = "A String!",
                Item1s = new List<Item1>(),
                Item2s = new List<Item2>()
            };

            var item1Sheet = new WorksheetConfiguration(typeof(Item1), "Item 1s", "Item1s", config.StylesheetProvider);
            config.AddWorksheet(item1Sheet);

            var item2Sheet = new WorksheetConfiguration(typeof(Item2), "Item 2s", "Item2s", config.StylesheetProvider);
            config.AddWorksheet(item2Sheet);

            using (var output = new MemoryStream())
            {
                config.Export(dataSource, output);
                var document = SpreadsheetDocument.Open(output, false);

                var sheets = document.WorkbookPart.Workbook.Sheets;

                Assert.AreEqual(2, sheets.Count());

                var firstSheet = sheets.ChildElements[0];
                var firstSheetNameAttrib = firstSheet.GetAttribute("name", null);
                Assert.IsNotNull(firstSheetNameAttrib);
                Assert.AreEqual("Item 1s", firstSheetNameAttrib.Value);

                var secondSheet = sheets.ChildElements[1];
                var secondSheetNameAttrib = secondSheet.GetAttribute("name", null);
                Assert.IsNotNull(secondSheetNameAttrib);
                Assert.AreEqual("Item 2s", secondSheetNameAttrib.Value);
            }
        }

        [TestMethod]
        public void Can_Export_Columns()
        {
            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));

            DateTime christmas2012 = new DateTime(2012, 12, 25, 11, 59, 30);
            DateTime newYears2012 = new DateTime(2013, 1, 1);

            var item1 = new SingleTableItem() { I = 0, ADate = christmas2012, StringField = "String 0" };
            var item2 = new SingleTableItem() { I = 1, ADate = newYears2012, StringField = "String 2" };

            // item 3 and 4 are here so we can be sure the null/empty values serialize out.
            // assertions are omitted because of the shared string table.
            var item3 = new SingleTableItem() { I = 2, ADate = newYears2012, StringField = null };
            var item4 = new SingleTableItem() { I = 2, ADate = newYears2012, StringField = string.Empty };

            var dataSource = new SingleTableHierarchy()
            {
                SingleTableItems = new List<SingleTableItem>() { item1, item2, item3, item4 }
            };

            var singleTableItemSheet = new WorksheetConfiguration(typeof(SingleTableItem), "Single Table", "SingleTableItems", config.StylesheetProvider);
            singleTableItemSheet.SheetName = "Item 1s";
            singleTableItemSheet.AddColumn("I", SingleTableHierarchyIMember);
            var dateColumn = singleTableItemSheet.AddColumn("A Date", SingleTableHierarchyADateMember);
            dateColumn.CellFormat = config.StylesheetProvider.DateFormat;
            singleTableItemSheet.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            config.AddWorksheet(singleTableItemSheet);

            using (var output = new MemoryStream())
            {
                config.Export(dataSource, output);
                var document = SpreadsheetDocument.Open(output, false);

                var sheets = document.WorkbookPart.Workbook.Sheets;

                Assert.AreEqual(1, sheets.Count());

                var firstSheet = sheets.ChildElements[0];
                var sheetData = document.WorkbookPart.WorksheetParts.First().Worksheet.GetFirstChild<SheetData>();

                Assert.IsNotNull(sheetData);

                var rows = sheetData.Elements<Row>().ToList();
                Assert.AreEqual(5, rows.Count);

                var firstDataRow = rows[1];
                var A2 = firstDataRow.Elements<Cell>().ToList()[0];
                Assert.AreEqual("0", A2.CellValue.Text, "A2");

                var B2 = firstDataRow.Elements<Cell>().ToList()[1];
                Assert.AreEqual(christmas2012.ToOADate().ToString(), B2.CellValue.Text, "B2");

                var C2 = firstDataRow.Elements<Cell>().ToList()[2];
                Assert.AreEqual("String 0", C2.CellValue.Text, "C2");

                var secondDataRow = rows[2];
                var A3 = secondDataRow.Elements<Cell>().ToList()[0];
                Assert.AreEqual("1", A3.CellValue.Text, "A3");

                var B3 = secondDataRow.Elements<Cell>().ToList()[1];
                Assert.AreEqual(newYears2012.ToOADate().ToString(), B3.CellValue.Text, "B3");

                var C3 = secondDataRow.Elements<Cell>().ToList()[2];
                Assert.AreEqual("String 2", C3.CellValue.Text, "C3");
            }
        }

        [TestMethod]
        public void Export_Null_In_Required_Column_Throws_Exception()
        {
            var config = new WorkbookConfiguration(typeof(SingleTableHierarchy));
            var dataSource = new SingleTableHierarchy()
            {
                SingleTableItems = new List<SingleTableItem>()
                {
                    new SingleTableItem{ StringField = null, I = 1 }
                }
            };

            var itemsSheetConfig = new WorksheetConfiguration(typeof(Item1), "Single Table Items", "SingleTableItems", config.StylesheetProvider);
            var columnConfig = itemsSheetConfig.AddColumn("String Field", SingleTableHierarchyStringFieldMember);
            columnConfig.Required = true;
            config.AddWorksheet(itemsSheetConfig);

            try
            {
                using (var output = new MemoryStream())
                {
                    config.Export(dataSource, output);
                    Assert.Fail("Should have failed due to exporting null to required column.");
                }
            }
            catch (NullableColumnViolationException ex)
            {
                Assert.AreEqual(
                    "Cannot export null value into non-nullable column \"String Field\" on sheet \"Single Table Items\".",
                    ex.Message);
            }
        }

        private static void SaveToFile(MemoryStream ms, string name)
        {
            using (var fs = File.OpenWrite(name))
                ms.WriteTo(fs);
        }

        private static void OpenInExcel(string excelFile)
        {
            System.Diagnostics.Process.Start(excelFile);
        }
    }
}
