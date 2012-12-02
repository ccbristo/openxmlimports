using System;
using System.Collections.Generic;

namespace ExcelImports
{
    internal class FuncEqualityComparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> equalityFunc;

        public FuncEqualityComparer(Func<T, T, bool> equalityFunc)
        {
            this.equalityFunc = equalityFunc;
        }

        public bool Equals(T x, T y)
        {
            return equalityFunc(x, y);
        }

        public int GetHashCode(T obj)
        {
            if (obj == null)
            {
                return 0;
            }

            return obj.GetHashCode();
        }
    }
}
