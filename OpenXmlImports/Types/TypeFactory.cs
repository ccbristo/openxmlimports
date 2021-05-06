using System;

namespace OpenXmlImports.Types
{
    public static class TypeFactory
    {
        public static IType GetType(Type type)
        {
            if (type.Is<bool>())
                return new BooleanType();

            if (type.Is<int>())
                return new IntType();

            if (type.Is<string>())
                return new StringType();

            if (type.Is<DateTime>())
                return new DateTimeType();

            if (type.Is<decimal>())
                return new DecimalType();

            if (type.Is<Enum>())
                return new EnumStringType(type);

            if (type.Is<byte>())
                return new ByteType();

            if (type.Is<sbyte>())
                return new SByteType();

            if (type.Is<short>())
                return new ShortType();

            if (type.Is<ushort>())
                return new UShortType();

            if (type.Is<uint>())
                return new UIntType();

            if (type.Is<long>())
                return new LongType();
            
            if (type.Is<ulong>())
                return new ULongType();

            if (type.Is<char>())
                return new CharType();

            if (type.Is<double>())
                return new DoubleType();

            if (type.Is<float>())
                return new SingleType();

            if (type.Is<Guid>())
                return new GuidType();

            throw new ArgumentOutOfRangeException("type",
                string.Format("Could not determine type for {0}", type.Name));
        }
    }
}
