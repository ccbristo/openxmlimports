using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlImports.Types;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class EnumTests
    {
        public class SourceWorkbook
        {
            public Color Color { get; set; }
            public Color? NullableColor { get; set; }
            public List<SourceSheet> Items { get; set; }
        }

        public class SourceSheet
        {
            public Color A { get; set; }
            public Color? B { get; set; }
        }

        public enum Color
        {
            Red,
            White,
            Blue,
            BurntSienna
        }

        [TestMethod]
        public void Can_Import_And_Export_Enums()
        {
            var config = OpenXmlConfiguration.Workbook<SourceWorkbook>()
                .Create();

            var source = new SourceWorkbook
            {
                Color = Color.White,
                NullableColor = Color.Red,
                Items = new List<SourceSheet>
                {
                    new SourceSheet{ A = Color.Red, B = null },
                    new SourceSheet{ A = Color.White, B = Color.Blue },
                    new SourceSheet{ A = Color.BurntSienna, B = Color.White }
                }
            };

            using (var output = new MemoryStream())
            {
                config.Export(source, output);
                output.Position = 0L;
                var imported = config.Import(output);

                Assert.AreEqual(Color.White, imported.Color, "imported.Color");
                Assert.AreEqual(Color.Red, imported.NullableColor, "imported.NullableColor");

                Assert.AreEqual(3, imported.Items.Count);

                Assert.AreEqual(Color.Red, imported.Items[0].A, "imported.Items[0].A");
                Assert.IsNull(imported.Items[0].B, "imported.Items[0].B");

                Assert.AreEqual(Color.White, imported.Items[1].A, "imported.Items[1].A");
                Assert.AreEqual(Color.Blue, imported.Items[1].B, "imported.Items[1].B");

                Assert.AreEqual(Color.BurntSienna, imported.Items[2].A, "imported.Items[2].A");
                Assert.AreEqual(Color.White, imported.Items[2].B, "imported.Items[2].B");
            }
        }

        [TestMethod]
        public void EnumsAreFormattedWithCamelCase()
        {
            var enumStringType = new EnumStringType(typeof(Color));
            var formatted = enumStringType.Format(Color.BurntSienna.ToString());
            Assert.AreEqual("Burnt Sienna", formatted);
        }

        [TestMethod]
        public void EnumsCanBeParsedExactly()
        {
            var enumStringType = new EnumStringType(typeof(Color));
            var parsed = enumStringType.Parse("BurntSienna");
            Assert.AreEqual(Color.BurntSienna, parsed);
        }

        [TestMethod]
        public void EnumsCanBeParsedByCamelCase()
        {
            var enumStringType = new EnumStringType(typeof(Color));
            var parsed = enumStringType.Parse("Burnt Sienna");
            Assert.AreEqual(Color.BurntSienna, parsed);
        }
    }
}