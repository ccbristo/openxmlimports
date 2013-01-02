using System;
using OpenXmlImports.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class ColumnReferenceTests
    {
        [TestMethod]
        public void Test_All_One_Letter_Cols()
        {
            var colRef = new ColumnReference();

            Assert.AreEqual(0, colRef.Value);
            Assert.AreEqual(string.Empty, colRef.ToString());
            colRef++;

            for (char c = 'A'; c <= 'Z'; c++)
            {
                Assert.AreEqual(c.ToString(), colRef.ToString());
                colRef++;
            }

            Assert.AreEqual("AA", colRef.ToString());
        }

        [TestMethod]
        public void Test_All_Two_Letter_Cols()
        {
            var colRef = new ColumnReference("AA");

            int currentValue = 27;

            for (char c1 = 'A'; c1 <= 'Z'; c1++)
            {
                for (char c2 = 'A'; c2 <= 'Z'; c2++)
                {
                    Assert.AreEqual(currentValue, colRef.Value);

                    string col = c1.ToString() + c2.ToString();
                    Assert.AreEqual(col, colRef.ToString());

                    colRef++;
                    currentValue++;
                }
            }

            Assert.AreEqual("AAA", colRef.ToString());
        }

        [TestMethod]
        public void Test_All_Three_Letter_Cols()
        {
            var colRef = new ColumnReference("AAA");

            int currentValue = 703;

            for (char c1 = 'A'; c1 <= 'Z'; c1++)
            {
                for (char c2 = 'A'; c2 <= 'Z'; c2++)
                {
                    for (char c3 = 'A'; c3 <= 'Z'; c3++)
                    {
                        Assert.AreEqual(currentValue, colRef.Value);

                        string col = c1.ToString() + c2.ToString() + c3.ToString();
                        Assert.AreEqual(col, colRef.ToString());

                        if (col.Equals("XFD")) // last usable column
                            return;

                        colRef++;
                        currentValue++;
                    }
                }
            }

        }

        [TestMethod]
        public void Test_Out_Of_Range_Column()
        {
            try
            {
                var outOfRange = new ColumnReference("XFE");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                Assert.AreEqual("column", ex.ParamName);
            }
        }
    }
}
