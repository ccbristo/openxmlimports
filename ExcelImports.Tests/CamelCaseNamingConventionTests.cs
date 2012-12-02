﻿using System;
using System.Linq.Expressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelImports.Tests
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
            internal decimal CapsINTRAWord { get; set; }
        }

        [TestMethod]
        public void TestSimpleField()
        {
            DoTest(x => x.SimpleField, "Simple Field");
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

        private void DoTest<T>(Expression<Func<HasMembersForTesting, T>> memberExp, string expectedName)
        {
            var member = memberExp.GetMemberInfo();

            CamelCaseNamingConvention namer = new CamelCaseNamingConvention();

            var name = namer.GetName(member);
            Assert.AreEqual(expectedName, name, false);
        }
    }
}