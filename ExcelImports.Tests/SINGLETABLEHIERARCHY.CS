using System;
using System.Collections.Generic;

namespace ExcelImports.Tests
{
    public class SingleTableHierarchy
    {
        public List<SingleTableItem> SingleTableItems { get; set; }
    }

    public class SingleTableItem
    {
        public string StringField;
        public int I { get; set; }
        public DateTime? ADate { get; set; }

    }
}
