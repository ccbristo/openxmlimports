using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
{
    [TestClass]
    public class ExcelConfigurationTests
    {
        [TestMethod]
        public void DefaultConfig()
        {
            var config = ExcelConfiguration.Workbook<SimpleHierarchy>()
                .Create();

            Assert.AreEqual(2, config.Count(), "number of worksheet configs");

            var item1Config = config.Single(wsc => wsc.BoundType == typeof(Item1));
            Assert.AreEqual(2, item1Config.Columns.Count());

            var item1NameColumn = item1Config.Single(column => StringComparer.Ordinal.Equals(column.Name, "Name"));
            Assert.AreEqual("Name", item1NameColumn.Member.Name);

            var item1ValueColumn = item1Config.Single(column => StringComparer.Ordinal.Equals(column.Name, "Value"));
            Assert.AreEqual(item1ValueColumn.Member.Name, "Value");

            var item2Config = config.Single(wsc => wsc.BoundType == typeof(Item2));
            Assert.AreEqual(2, item2Config.Columns.Count());

            var item2TitleColumn = item2Config.Single(column => StringComparer.Ordinal.Equals(column.Name, "Title"));
            Assert.AreEqual(item2TitleColumn.Member.Name, "Title");

            var item2ADateColumn = item2Config.Single(column => StringComparer.Ordinal.Equals(column.Name, "A Date"));
            Assert.AreEqual(item2ADateColumn.Member.Name, "ADate");
        }

        [TestMethod]
        public void Override_Sheet_Name()
        {
            var config = ExcelConfiguration.Workbook<SimpleHierarchy>()
                .Worksheet(sh => sh.Item1s, item1Sheet =>
                    {
                        item1Sheet.Named("The Item1s");
                    })
                .Create();

            Assert.AreEqual(2, config.Count(), "number of worksheet configs");

            var item1Config = config.Single(wsc => wsc.BoundType == typeof(Item1));
            Assert.AreEqual("The Item1s", item1Config.SheetName);
        }

        [TestMethod]
        public void Override_Column_Name()
        {
            var config = ExcelConfiguration.Workbook<SimpleHierarchy>()
                .Worksheet(sh => sh.Item1s, item1Sheet =>
                {
                    item1Sheet.Column(item => item.Name, nameCol =>
                    {
                        nameCol.Named("The Name Column");
                    });

                })
                .Create();

            Assert.AreEqual(2, config.Count(), "number of worksheet configs");

            var item1Config = config.Single(wsc => wsc.BoundType == typeof(Item1));
            var nameColumn = item1Config.Single(col => StringComparer.Ordinal.Equals("Name", col.Member.Name));
            Assert.AreEqual("The Name Column", nameColumn.Name);
        }

        [TestMethod]
        public void AAAAAA()
        {
            var workbook = new SimpleHierarchy
            {
                I = 5,
                S = "Some text",
                Item1s = new List<Item1>()
                {
                    new Item1{ Name = "First Item 1", Value = 5 },
                    new Item1{ Name = "Second Item 1", Value = 4 },
                    new Item1{ Name = "Third Item 1", Value = 3 },
                    new Item1{ Name = "Fourth Item 1", Value = 2 },
                    new Item1{ Name = "Fifth Item 1", Value = 1 }
                },
                Item2s = new List<Item2>()
                {
                    new Item2{ Title = "First Item 2", ADate = new DateTime(1981, 8, 28) },
                    new Item2{ Title = "Second Item 2", ADate = new DateTime(1982, 8, 27) },
                    new Item2{ Title = "Third Item 2", ADate = new DateTime(1953, 7, 11) },
                    new Item2{ Title = "Fourth Item 2", ADate = new DateTime(1954, 2, 22) },
                    new Item2{ Title = "Fifth Item 2", ADate = new DateTime(1984, 3, 7) }
                }
            };

            var config = ExcelConfiguration.Workbook<SimpleHierarchy>()
                .Create();

            var dateColumn = config.GetWorksheet("Item 2s").Single(c => c.Name == "A Date");
            Assert.AreSame(config.StylesheetProvider.DateFormat, dateColumn.CellFormat);

            using (var fs = System.IO.File.Open("output.xlsx", System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite))
                config.Export(workbook, fs);

            System.Diagnostics.Process.Start("output.xlsx");
        }

        //[TestMethod]
        public void Test1()
        {
            Assert.Fail("So far, all this test does is check some of the syntax I would like to have.");

            ExcelConfiguration.Workbook<PropertyRateSet>()
                .Worksheet("Instructions", instructionsSheet => instructionsSheet.ExportOnly())
                .Worksheet(prs => prs.StateGroups, stateGroupSheet =>
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