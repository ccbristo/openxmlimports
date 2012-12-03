using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelImports.Core
{
    public class ColumnConfiguration
    {
    }

    public class ColumnConfiguration<T>
        : ColumnConfiguration
    {
        private bool listValidOptions = false;
        private IEnumerable options;
        private IEqualityComparer<T> Comparer = EqualityComparer<T>.Default;

        public ColumnConfiguration<T> OneOf(IEnumerable<T> options)
        {
            this.options = options;
            return this;
        }

        public ColumnConfiguration<T> OneOf<T2>(IEnumerable<T2> options, Func<T, T2, bool> comparison)
        {
            this.options = options;
            return this;
        }

        public ColumnConfiguration<T> ListValidValuesOnError(bool listValidOptions)
        {
            this.listValidOptions = listValidOptions;
            return this;
        }
    }
}
