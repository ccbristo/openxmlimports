﻿using System;
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
            public object Has_Under_Scores { get; set; }
        }

        [TestMethod]
        public void TestSimpleField()
        {
            DoTest(x => x.SimpleField, "Simple Field");
        }

        [TestMethod]
        public void TestADate()
        {
            DoTest(x => x.ADate, "A Date");
        }

        [TestMethod]
        public void TestSimpleProperty()
        {
            DoTest(x => x.SimpleProperty, "Simple Property");
        }

        [TestMethod]
        public void TestCAPPrefix()
        {
            DoTest(x => x.CAPPrefix, "CAP Prefix");
        }

        [TestMethod]
        public void TestSuffixCAP()
        {
            DoTest(x => x.SuffixCAP, "Suffix CAP");
        }

        [TestMethod]
        public void TestCapsINTRAWord()
        {
            DoTest(x => x.CapsINTRAWord, "Caps INTRA Word");
        }

        [TestMethod]
        public void TestNumbers()
        {
            DoTest(x => x.Has150Numbers, "Has 150 Numbers");
        }

        [TestMethod]
        public void TestUnderscores()
        {
            DoTest(x => x.Has_Under_Scores, "Has Under Scores");
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