using System;
using DocumentFormat.OpenXml;

namespace OpenXmlImports.Core
{
    [System.Diagnostics.DebuggerDisplay("{ToString()}")]
    public struct ColumnReference
    {
        public static readonly ColumnReference Empty = new ColumnReference();
        public static readonly ColumnReference MinValue = new ColumnReference(1); // column A
        public static readonly ColumnReference MaxValue = new ColumnReference(16384); // column XFD

        readonly int mValue;
        public int Value
        {
            get { return mValue; }
        }

        public ColumnReference(string column)
        {
            mValue = Parse(column);

            if (mValue < MinValue.Value || mValue > MaxValue.Value)
                throw new ArgumentOutOfRangeException("column");
        }

        private ColumnReference(int column)
        {
            mValue = column;
        }

        private static int Parse(string column)
        {
            int result = 0;
            string col = column.ToUpper();
            int shift = 1;

            for (int i = col.Length - 1; i >= 0; i--)
            {
                char letter = col[i];
                int colNum = letter - 'A' + 1;

                result += colNum * shift;
                shift *= 26;
            }

            return result;
        }

        public static ColumnReference operator ++(ColumnReference colRef)
        {
            return new ColumnReference(colRef.Value + 1);
        }

        public static ColumnReference operator --(ColumnReference colRef)
        {
            return new ColumnReference(colRef.Value - 1);
        }

        public static bool operator >(ColumnReference a, ColumnReference b)
        {
            return a.Value > b.Value;
        }

        public static bool operator >=(ColumnReference a, ColumnReference b)
        {
            return a.Value >= b.Value;
        }

        public static bool operator <(ColumnReference a, ColumnReference b)
        {
            return a.Value < b.Value;
        }

        public static bool operator <=(ColumnReference a, ColumnReference b)
        {
            return a.Value <= b.Value;
        }

        public static bool operator ==(ColumnReference a, ColumnReference b)
        {
            return a.Value == b.Value;
        }

        public static bool operator !=(ColumnReference a, ColumnReference b)
        {
            return a.Value != b.Value;
        }

        public static implicit operator StringValue(ColumnReference colRef)
        {
            return new StringValue(colRef.ToString());
        }

        public static implicit operator string(ColumnReference colRef)
        {
            return colRef.ToString();
        }

        public static implicit operator ColumnReference(string column)
        {
            return new ColumnReference(column);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ColumnReference))
                return false;

            return ((ColumnReference)obj).Value == this.Value;
        }

        public override int GetHashCode()
        {
            return Value;
        }

        public override string ToString()
        {
            string columnString = string.Empty;
            decimal columnNumber = Value;

            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 'A');
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            return columnString;
        }
    }
}
