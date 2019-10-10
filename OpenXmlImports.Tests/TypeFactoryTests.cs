using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXmlImports.Types;

namespace OpenXmlImports.Tests
{
    [TestClass]
    public class TypeFactoryTests
    {
        [TestMethod]
        public void CanGetTypeForBool()
        {
            Verify<bool, BooleanType>();
        }

        [TestMethod]
        public void CanGetTypeForInt()
        {
            Verify<int, IntType>();
        }

        [TestMethod]
        public void CanGetTypeForString()
        {
            Verify<string, StringType>();
        }

        [TestMethod]
        public void CanGetTypeForDateTime()
        {
            Verify<DateTime, DateTimeType>();
        }

        [TestMethod]
        public void CanGetTypeForDecimal()
        {
            Verify<decimal, DecimalType>();
        }

        private enum E { A, B }

        [TestMethod]
        public void CanGetTypeForByte()
        {
            Verify<byte, ByteType>();
        }

        [TestMethod]
        public void CanGetTypeForSByte()
        {
            Verify<sbyte, SByteType>();
        }

        [TestMethod]
        public void CanGetTypeForShort()
        {
            Verify<short, ShortType>();
        }

        [TestMethod]
        public void CanGetTypeForUShort()
        {
            Verify<ushort, UShortType>();
        }

        [TestMethod]
        public void CanGetTypeForUInt()
        {
            Verify<uint, UIntType>();
        }

        [TestMethod]
        public void CanGetTypeForULong()
        {
            Verify<ulong, ULongType>();
        }

        [TestMethod]
        public void CanGetTypeForChar()
        {
            Verify<char, CharType>();
        }

        [TestMethod]
        public void CanGetTypeForDouble()
        {
            Verify<double, DoubleType>();
        }

        [TestMethod]
        public void CanGetTypeForFloat()
        {
            Verify<float, SingleType>();
        }

        [TestMethod]
        public void CanGetTypeForGuid()
        {
            Verify<Guid, GuidType>();
        }

        private void Verify<TStorage, TApplication>()
        {
            var applicationType = TypeFactory.GetType(typeof(TStorage));
            Assert.IsNotNull(applicationType);
            Assert.IsTrue(applicationType.GetType() == typeof(TApplication));
        }
    }
}
