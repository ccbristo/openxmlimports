using System.Collections.Generic;

namespace OpenXmlImports.Tests
{
    public class InheritanceHierarchy
    {
        public List<DerivedClass> Items { get; set; }
    }

    public abstract class BaseClass
    {
        public int OnBaseClass { get; set; }
    }

    public class DerivedClass : BaseClass
    {
        public int OnDerivedClass { get; set; }
    }
}