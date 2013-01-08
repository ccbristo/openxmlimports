
using System;
using System.Linq.Expressions;
using System.Reflection;
namespace OpenXmlImports.Tests
{
    public class RootPropertiesHierarchy
    {
        public Item SingleItem { get; set; }
        public int I { get; set; }
        public string Data { get; set; }
    }

    public class Item
    {
        public int J { get; set; }
        public string S { get; set; }
    }

    public static class RootPropertiesHierarchyMembers
    {
        public static readonly MemberInfo Data;
        public static readonly MemberInfo I;
        public static readonly MemberInfo SingleItem;

        public static class SingleItemMembers
        {
            public static readonly MemberInfo J;
            public static readonly MemberInfo S;

            static SingleItemMembers()
            {
                Expression<Func<Item, int>> jExp = i => i.J;
                J = jExp.GetMemberInfo();

                Expression<Func<Item, string>> sExp = i => i.S;
                S = sExp.GetMemberInfo();
            }
        }

        static RootPropertiesHierarchyMembers()
        {
            Expression<Func<RootPropertiesHierarchy, string>> dataExp = r => r.Data;
            Data = dataExp.GetMemberInfo();

            Expression<Func<RootPropertiesHierarchy, int>> iExp = r => r.I;
            I = iExp.GetMemberInfo();

            Expression<Func<RootPropertiesHierarchy, Item>> singleItemExp = r => r.SingleItem;
            SingleItem = singleItemExp.GetMemberInfo();

            RootPropertiesHierarchy x = new RootPropertiesHierarchy
            {
                Data = "test",
                I = 5,
                SingleItem = new Item { J = 1, S = "" }
            };
        }
    }
}
