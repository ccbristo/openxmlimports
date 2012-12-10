using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelImports
{
    public class ColumnBuilder
    {

    }

    public class ColumnBuilder<TColumn>
        : ColumnBuilder
    {
        private bool listValidOptions = false;
        private IEnumerable options;
        private IEqualityComparer<TColumn> Comparer = EqualityComparer<TColumn>.Default;

        public ColumnBuilder<TColumn> OneOf(IEnumerable<TColumn> options)
        {
            this.options = options;
            return this;
        }

        public ColumnBuilder<TColumn> OneOf<TOption>(IEnumerable<TOption> options, Func<TColumn, TOption, bool> comparison)
        {
            this.options = options;
            return this;
        }

        public ColumnBuilder<TColumn> ListValidValuesOnError(bool listValidOptions)
        {
            this.listValidOptions = listValidOptions;
            return this;
        }
    }
}
