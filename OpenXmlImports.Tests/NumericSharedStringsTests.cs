using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class NumericSharedStringsTests
    {
        private class Workbook
        {
            public List<Item> Items { get; set; }
        }
        
        private class Item
        {
            public long LongNumber { get; set; }
            public string TextNumber { get; set; }
            public byte Field1 { get; set; }
            public decimal  Field2 { get; set; }
            public double Field3 { get; set; }
            public int Field4 { get; set; }
            public sbyte Field5 { get; set; }
            public short Field6 { get; set; }
            public float Field7 { get; set; }
            public uint Field8 { get; set; }
            public ulong Field9 { get; set; }
            public ushort Field10 { get; set; }
        }

        [TestMethod]
        public void NumericSharedStringsAreImportedCorrectly()
        {
            var config = OpenXmlConfiguration.Workbook<Workbook>()
                .Create();

            using (var reader = typeof(NumericSharedStringsTests).Assembly.GetManifestResourceStream("OpenXmlImports.Tests.TestFiles.Numeric_Shared_Strings.xlsx"))
            {
                var workbook = config.Import(reader);
                
                Assert.AreEqual(1, workbook.Items.Count);

                var item = workbook.Items[0];
                Assert.AreEqual(161021544791149119, item.LongNumber);
                Assert.AreEqual("600001826", item.TextNumber);
                Assert.AreEqual(1, item.Field1);
                Assert.AreEqual(2, item.Field2);
                Assert.AreEqual(3, item.Field3);
                Assert.AreEqual(4, item.Field4);
                Assert.AreEqual(5, item.Field5);
                Assert.AreEqual(6, item.Field6);
                Assert.AreEqual(7, item.Field7);
                Assert.AreEqual(8U, item.Field8);
                Assert.AreEqual(9UL, item.Field9);
                Assert.AreEqual((ushort)10, item.Field10);
            }
        }
    }
}