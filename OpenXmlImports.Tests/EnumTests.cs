using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            Blue
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
                    new SourceSheet{ A = Color.Blue, B = Color.White }
                }
            };

            using (var fs = System.IO.File.Create("out.xlsx"))
            {
                config.Export(source, fs);
                fs.Position = 0L;
                var imported = (SourceWorkbook)config.Import(fs);

                Assert.AreEqual(Color.White, imported.Color, "imported.Color");
                Assert.AreEqual(Color.Red, imported.NullableColor, "imported.NullableColor");

                Assert.AreEqual(3, imported.Items.Count);

                Assert.AreEqual(Color.Red, imported.Items[0].A, "imported.Items[0].A");
                Assert.IsNull(imported.Items[0].B, "imported.Items[0].B");

                Assert.AreEqual(Color.White, imported.Items[1].A, "imported.Items[1].A");
                Assert.AreEqual(Color.Blue, imported.Items[1].B, "imported.Items[1].B");

                Assert.AreEqual(Color.Blue, imported.Items[2].A, "imported.Items[2].A");
                Assert.AreEqual(Color.White, imported.Items[2].B, "imported.Items[2].B");
            }
        }
    }
}
