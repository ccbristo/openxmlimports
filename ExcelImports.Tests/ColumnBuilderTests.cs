using System;
using System.Linq.Expressions;
using ExcelImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
{
    [TestClass]
    public class ColumnBuilderTests
    {
        class X
        {
            public DateTime ADate { get; set; }
            public DateTime? NullableDate { get; set; }
        }

        [TestMethod]
        public void Should_Default_Date_Times_To_Date_Format()
        {
            Expression<Func<X, DateTime>> dateExp = x => x.ADate;
            var stylesheetProvider = new DefaultStylesheetProvider();
            var column = new ColumnBuilder<DateTime>("A Date", dateExp.GetMemberInfo(), stylesheetProvider);

            Assert.AreSame(stylesheetProvider.DateFormat, column.Configuration.CellFormat, "Wrong cell format applied for DateTime.");
        }

        [TestMethod]
        public void Should_Default_Nullable_Date_Times_To_Date_Format()
        {
            Expression<Func<X, DateTime?>> nullableDateExp = x => x.NullableDate;
            var stylesheetProvider = new DefaultStylesheetProvider();
            var column = new ColumnBuilder<DateTime>("A Date", nullableDateExp.GetMemberInfo(), stylesheetProvider);

            Assert.AreSame(stylesheetProvider.DateFormat, column.Configuration.CellFormat, "Wrong cell format applied for DateTime?.");
        }
    }
}
