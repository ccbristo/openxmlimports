using System;
using System.Linq.Expressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class CamelCaseNamingConventionTests
    {
        private class HasMembersForTesting
        {
            public string SimpleProperty { get; set; }
            public string SimpleField = null;
            public int CAPPrefix { get; set; }
            public object SuffixCAP { get; set; }
            public decimal CapsINTRAWord { get; set; }
            public int Has150Numbers { get; set; }
            public string ADate { get; set; }
            public object has_under_scores { get; set; }
            public object Under_Scores_With_Camel_Case { get; set; }
        }

        [TestMethod]
        public void SimpleField()
        {
            DoTest(x => x.SimpleField, "Simple Field");
        }

        [TestMethod]
        public void ADate()
        {
            DoTest(x => x.ADate, "A Date");
        }

        [TestMethod]
        public void SimpleProperty()
        {
            DoTest(x => x.SimpleProperty, "Simple Property");
        }

        [TestMethod]
        public void CAPPrefix()
        {
            DoTest(x => x.CAPPrefix, "CAP Prefix");
        }

        [TestMethod]
        public void SuffixCAP()
        {
            DoTest(x => x.SuffixCAP, "Suffix CAP");
        }

        [TestMethod]
        public void CapsINTRAWord()
        {
            DoTest(x => x.CapsINTRAWord, "Caps INTRA Word");
        }

        [TestMethod]
        public void Has150Numbers()
        {
            DoTest(x => x.Has150Numbers, "Has 150 Numbers");
        }

        [TestMethod]
        public void has_under_scores()
        {
            DoTest(x => x.has_under_scores, "has under scores");
        }

        [TestMethod]
        public void Under_Scores_With_Camel_Case()
        {
            DoTest(x => x.Under_Scores_With_Camel_Case, "Under Scores With Camel Case");
        }

        private void DoTest<T>(Expression<Func<HasMembersForTesting, T>> memberExp, string expectedName)
        {
            var member = memberExp.GetMemberInfo();

            CamelCaseNamingConvention namer = new CamelCaseNamingConvention();

            var name = namer.GetName(member);
            Assert.AreEqual(expectedName, name, false);
        }
    }
}
