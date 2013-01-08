using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    public class NullabilityWorkbookEntity
    {
        public IList<NullabilityWorksheetEntity> Items { get; set; }
    }

    public class NullabilityWorksheetEntity
    {
        public int Value { get; set; }
        public string Reference { get; set; }
        public int? NullableProp { get; set; }
    }

    [TestClass]
    public class OpenXmlConfigurationTests
    {
        [TestMethod]
        public void DefaultConfig()
        {
            var config = OpenXmlConfiguration.Workbook<SimpleHierarchy>()
                .Create();

            Assert.AreEqual(2, config.Worksheets.Count(), "number of worksheet configs");

            var item1Config = config.Worksheets.Single(wsc => wsc.BoundType == typeof(Item1));
            Assert.AreEqual(2, item1Config.Columns.Count());

            var item1NameColumn = item1Config.Columns.Single(column => StringComparer.Ordinal.Equals(column.Name, "Name"));
            Assert.AreEqual("Name", item1NameColumn.Member.Name);

            var item1ValueColumn = item1Config.Columns.Single(column => StringComparer.Ordinal.Equals(column.Name, "Value"));
            Assert.AreEqual(item1ValueColumn.Member.Name, "Value");

            var item2Config = config.Worksheets.Single(wsc => wsc.BoundType == typeof(Item2));
            Assert.AreEqual(2, item2Config.Columns.Count());

            var item2TitleColumn = item2Config.Columns.Single(column => StringComparer.Ordinal.Equals(column.Name, "Title"));
            Assert.AreEqual(item2TitleColumn.Member.Name, "Title");

            var item2ADateColumn = item2Config.Columns.Single(column => StringComparer.Ordinal.Equals(column.Name, "A Date"));
            Assert.AreEqual(item2ADateColumn.Member.Name, "ADate");
        }

        [TestMethod]
        public void Override_Sheet_Name()
        {
            var config = OpenXmlConfiguration.Workbook<SimpleHierarchy>()
                .List(sh => sh.Item1s, (item1Sheet, styles) =>
                    {
                        item1Sheet.Named("The Item1s");
                    })
                .Create();

            Assert.AreEqual(2, config.Worksheets.Count(), "number of worksheet configs");

            var item1Config = config.Worksheets.Single(wsc => wsc.BoundType == typeof(Item1));
            Assert.AreEqual("The Item1s", item1Config.SheetName);
        }

        [TestMethod]
        public void Override_Column_Name()
        {
            var config = OpenXmlConfiguration.Workbook<SimpleHierarchy>()
                .List(sh => sh.Item1s, (item1Sheet, styles) =>
                {
                    item1Sheet.Column(item => item.Name, nameCol =>
                    {
                        nameCol.Named("The Name Column");
                    });

                })
                .Create();

            Assert.AreEqual(2, config.Worksheets.Count(), "number of worksheet configs");

            var item1Config = config.Worksheets.Single(wsc => wsc.BoundType == typeof(Item1));
            var nameColumn = item1Config.Columns.Single(col => StringComparer.Ordinal.Equals("Name", col.Member.Name));
            Assert.AreEqual("The Name Column", nameColumn.Name);

            Assert.AreEqual(2, item1Config.Columns.Count(), "number of columns");
        }

        [TestMethod]
        public void Reference_Types_Should_Not_Be_Required_By_Default()
        {
            var config = OpenXmlConfiguration.Workbook<NullabilityWorkbookEntity>()
                .Create();

            var sheetConfig = config.Worksheets.Single(wsc => wsc.BoundType == typeof(NullabilityWorksheetEntity));
            var referenceColumn = sheetConfig.Columns.Single(col => StringComparer.Ordinal.Equals("Reference", col.Member.Name));

            Assert.IsFalse(referenceColumn.Required, "Reference column is required.");
        }

        [TestMethod]
        public void Values_Types_Should_Be_Required_By_Default()
        {
            var config = OpenXmlConfiguration.Workbook<NullabilityWorkbookEntity>()
                .Create();

            var sheetConfig = config.Worksheets.Single(wsc => wsc.BoundType == typeof(NullabilityWorksheetEntity));
            var valueColumn = sheetConfig.Columns.Single(col => StringComparer.Ordinal.Equals("Value", col.Member.Name));

            Assert.IsTrue(valueColumn.Required, "Value column is not required.");
        }

        [TestMethod]
        public void Explicitly_Require_Reference_Type()
        {
            var config = OpenXmlConfiguration.Workbook<NullabilityWorkbookEntity>()
                .List(sh => sh.Items, (itemSheet, styles) =>
                {
                    itemSheet.Column(item => item.Reference, refCol =>
                    {
                        refCol.Required(true);
                    });
                })
                .Create();

            var itemsConfig = config.Worksheets.Single(wsc => wsc.BoundType == typeof(NullabilityWorksheetEntity));
            var referenceColumn = itemsConfig.Columns.Single(col => StringComparer.Ordinal.Equals("Reference", col.Member.Name));

            Assert.IsTrue(referenceColumn.Required, "Reference column is not required.");
        }

        [TestMethod]
        public void Explicitly_Do_Not_Require_Value_Type()
        {
            var config = OpenXmlConfiguration.Workbook<NullabilityWorkbookEntity>()
                .List(sh => sh.Items, (itemSheet, styles) =>
                {
                    itemSheet.Column(item => item.Reference, refCol =>
                    {
                        refCol.Required(false);
                    });
                })
                .Create();

            var itemsConfig = config.Worksheets.Single(wsc => wsc.BoundType == typeof(NullabilityWorksheetEntity));
            var referenceColumn = itemsConfig.Columns.Single(col => StringComparer.Ordinal.Equals("Reference", col.Member.Name));

            Assert.IsFalse(referenceColumn.Required, "Reference column is required.");
        }

        [TestMethod]
        public void Root_Properties_And_Singletons_Are_Inferred()
        {
            var config = OpenXmlConfiguration.Workbook<RootPropertiesHierarchy>()
                .Create();

            RootPropertiesHierarchy hierarchy = new RootPropertiesHierarchy
                {
                    Data = "Some data.",
                    I = 3,
                    SingleItem = new Item { J = 4, S = "S" }
                };

            var detailsSheet = config.GetWorksheet("Details");

            Assert.AreEqual(2, detailsSheet.Columns.Count());
            Assert.AreEqual(2, config.Worksheets.Count(), "Should be 2 sheets in config.");
        }

        [TestMethod]
        public void Explicitly_Configure_Root_Properties()
        {
            var config = OpenXmlConfiguration.Workbook<RootPropertiesHierarchy>()
                .Singleton(rph => rph.SingleItem, (sheet, style) =>
                    {
                        sheet
                            .Named("Single Item Sheet")
                            .Column(item => item.J, jColumn =>
                            {
                                jColumn.Named("J Column");
                            });
                    })
                .RootProperties((sheet, style) =>
                    {
                        sheet.Named("Details Sheet")
                            .Column(rhp => rhp.Data, dataColumn =>
                            {
                                dataColumn.Named("Data Column");
                            });
                    }
                )
                .Create();

            Assert.AreEqual(2, config.Worksheets.Count(), "Worksheet count");

            var detailsSheet = config.Worksheets.SingleOrDefault(ws => ws.SheetName.Equals("Details Sheet"));
            Assert.IsNotNull(detailsSheet);

            Assert.AreEqual(2, detailsSheet.Columns.Count(), "detailsSheet.Columns.Count()");
            var dataCol = detailsSheet.Columns.SingleOrDefault(c => c.Name.Equals("Data Column"));
            Assert.IsNotNull(dataCol);

            var singleItemSheet = config.Worksheets.SingleOrDefault(ws => ws.SheetName.Equals("Single Item Sheet"));
            Assert.IsNotNull(singleItemSheet);

            Assert.AreEqual(2, singleItemSheet.Columns.Count(), "singleItemSheet.Columns.Count()");
            var jCol = singleItemSheet.Columns.SingleOrDefault(c => c.Name.Equals("J Column"));
            Assert.IsNotNull(jCol);
        }

        //[TestMethod]
        public void Test1()
        {
            Assert.Fail("So far, all this test does is check some of the syntax I would like to have.");

            OpenXmlConfiguration.Workbook<PropertyRateSet>()
                .Worksheet("Instructions", instructionsSheet => instructionsSheet.ExportOnly())
                .List(prs => prs.StateGroups, (stateGroupSheet, styles) =>
                    {
                        stateGroupSheet
                            .Column(sga => sga.State, columnConfig =>
                                {
                                    //columnConfig
                                    //    .OneOf(AllStates())
                                    //    .ListValidValuesOnError(false);
                                });

                    })
                // .Worksheet<SomeStuffWithLimits>(limitSheet =>
                //     {
                //         limitSheet.Named("Limits")
                //             .Column(
                //                 limitThing => new { limitThing.Occurrence, limitThing.Aggregate },
                //                 config =>
                //                 {
                //                     config.OneOf(LimitSet.All,
                //                         (anon, limit) => anon.Occurrence == limit.Occurrence &&
                //                             anon.Aggregate == limit.Aggregate);
                //                 });

               //     })
               ;
        }

        private IEnumerable<string> AllStates()
        {
            yield return "AK";
            yield return "AL";
            yield return "AR";
            yield return "AZ";
            yield return "CA";
            yield return "CO";
            yield return "CT";
            yield return "DC";
            yield return "DE";
            yield return "FL";
            yield return "GA";
            yield return "HI";
            yield return "IA";
            yield return "ID";
            yield return "IL";
            yield return "IN";
            yield return "KS";
            yield return "KY";
            yield return "LA";
            yield return "MA";
            yield return "MD";
            yield return "ME";
            yield return "MI";
            yield return "MN";
            yield return "MO";
            yield return "MS";
            yield return "MT";
            yield return "NC";
            yield return "ND";
            yield return "NE";
            yield return "NH";
            yield return "NJ";
            yield return "NM";
            yield return "NV";
            yield return "NY";
            yield return "OH";
            yield return "OK";
            yield return "OR";
            yield return "PA";
            yield return "PR";
            yield return "RI";
            yield return "SC";
            yield return "SD";
            yield return "TN";
            yield return "TX";
            yield return "UT";
            yield return "VA";
            yield return "VT";
            yield return "WA";
            yield return "WI";
            yield return "WV";
            yield return "WY";
        }
    }
}