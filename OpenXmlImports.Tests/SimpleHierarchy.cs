using System;
using System.Collections.Generic;

namespace OpenXmlImports.Tests
{
    public class SimpleHierarchy
    {
        public List<Item1> Item1s;
        public List<Item2> Item2s;
    }

    public class Item1
    {
        public string Name { get; set; }
        public int Value { get; set; }
    }

    public class Item2
    {
        public string Title { get; set; }
        public DateTime ADate { get; set; }
    }
}
