using System;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Core;

namespace OpenXmlImports
{
    public class ColumnBuilder
    {
        public ColumnConfiguration Configuration { get; private set; }

        internal ColumnBuilder()
        {
            this.Configuration = new ColumnConfiguration();
        }

        public ColumnBuilder Named(string name)
        {
            Configuration.Name = name;
            return this;
        }

        public ColumnBuilder Nullable()
        {
            return Nullable(true);
        }

        public ColumnBuilder Nullable(bool allowNulls)
        {
            Configuration.AllowNull = allowNulls;
            return this;
        }
    }

    public class ColumnBuilder<TColumn>
        : ColumnBuilder
    {
        // this ctor is referenced via reflection in WorksheetBuilder.ColumnFor
        public ColumnBuilder(string columnName, MemberInfo member, IStylesheetProvider stylesheet)
        {
            this.Named(columnName);
            this.Configuration.Member = member;

            var memberType = member.GetPropertyOrFieldType();
            if (memberType.IsValueType && !memberType.IsNullable())
                Configuration.AllowNull = false;

            // default datetimes to the date format
            if (typeof(TColumn).In(typeof(DateTime), typeof(DateTime?)))
                this.Configuration.CellFormat = stylesheet.DateFormat;
        }

        public new ColumnBuilder<TColumn> Named(string name)
        {
            base.Named(name);
            return this;
        }

        public ColumnBuilder<TColumn> Format(NumberingFormat format)
        {
            base.Configuration.CellFormat = format;
            return this;
        }

        public new ColumnBuilder<TColumn> Nullable()
        {
            base.Nullable();
            return this;
        }

        public new ColumnBuilder<TColumn> Nullable(bool allowNulls)
        {
            base.Nullable(allowNulls);
            return this;
        }

        // TODO [ccb] Implement these features.
        //public ColumnBuilder<TColumn> OneOf(IEnumerable<TColumn> options)
        //{
        //    return OneOf(options, EqualityComparer<TColumn>.Default)
        //}

        //// TODO [ccb] Bring back FuncEqualityComparer<TColumn, TOption>
        //public ColumnBuilder<TColumn> OneOf<TOption>(IEnumerable<TOption> options, Func<TColumn, TOption, bool> comparison)
        //{
        //    this.options = options;
        //    return this;
        //}

        //public ColumnBuilder<TColumn> OneOf(IEnumerable<TColumn> options, IEqualityComparer<TColumn> comparer)
        //{

        //}

        //public ColumnBuilder<TColumn> ListValidValuesOnError()
        //{
        //    return ListValidValuesOnError(true);
        //}

        //public ColumnBuilder<TColumn> ListValidValuesOnError(bool listValidOptions)
        //{
        //    Configuration.ListValidValuesOnError = listValidOptions;
        //    return this;
        //}
    }
}
