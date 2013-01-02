using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OpenXmlImports.Core;

namespace OpenXmlImports
{
    public class WorksheetBuilder
    {
        protected internal WorksheetConfiguration Configuration { get; set; }
        protected INamingConvention ColumnNamingConvention { get; set; }
        readonly IDictionary<MemberInfo, ColumnBuilder> Columns = new Dictionary<MemberInfo, ColumnBuilder>();

        public WorksheetBuilder(IStylesheetProvider stylesheetProvider)
            : this(new WorksheetConfiguration(stylesheetProvider))
        { }

        internal WorksheetBuilder(Type boundType, IStylesheetProvider stylesheetProvider)
            : this(new WorksheetConfiguration(boundType, stylesheetProvider))
        { }

        private WorksheetBuilder(WorksheetConfiguration config)
        {
            this.Configuration = config;
            this.ColumnNamingConvention = new CamelCaseNamingConvention();
        }

        public virtual WorksheetBuilder ExportOnly()
        {
            return ExportOnly(true);
        }

        public virtual WorksheetBuilder ExportOnly(bool exportOnly)
        {
            this.Configuration.ExportOnly = exportOnly;
            return this;
        }

        public WorksheetBuilder Named(string name)
        {
            this.Configuration.SheetName = name;
            return this;
        }

        protected void AddColumn(MemberInfo member, ColumnBuilder builder)
        {
            Columns[member] = builder;
        }

        internal void ConfigureColumns()
        {
            var columnMembers = Configuration.BoundType.GetMembers(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.IsPropertyOrField())
                .ToList();

            foreach (var member in columnMembers)
            {
                ColumnBuilder columnBuilder;
                if (Columns.TryGetValue(member, out columnBuilder))
                {
                    Configuration.AddColumn(columnBuilder.Configuration);
                    continue;
                }

                columnBuilder = ColumnFor(member);
                Columns[member] = columnBuilder;
                Configuration.AddColumn(columnBuilder.Configuration);
            }
        }

        private ColumnBuilder ColumnFor(MemberInfo member)
        {
            // since we don't know TColumn at compile type here, we have to construct it
            // via reflection.
            var ctor = typeof(ColumnBuilder<>).MakeGenericType(member.GetPropertyOrFieldType())
                .GetConstructor(new[] { typeof(string), typeof(MemberInfo), typeof(IStylesheetProvider) });

            string name = ColumnNamingConvention.GetName(member);
            var args = new object[] { name, member, Configuration.StylesheetProvider };
            return (ColumnBuilder)ctor.Invoke(args);
        }
    }

    public class WorksheetBuilder<T>
        : WorksheetBuilder
    {
        public WorksheetBuilder(string name, string memberName, IStylesheetProvider stylesheetProvider)
            : base(typeof(T), stylesheetProvider)
        {
            this.Named(name);
            this.Configuration.MemberName = memberName;
        }

        public new WorksheetBuilder<T> Named(string name)
        {
            base.Named(name);
            return this;
        }

        public WorksheetBuilder<T> Column<TColumn>(Expression<Func<T, TColumn>> columnExp,
            Action<ColumnBuilder<TColumn>> action)
        {
            var member = columnExp.GetMemberInfo();
            string columnName = ColumnNamingConvention.GetName(member);
            var columnBuilder = new ColumnBuilder<TColumn>(columnName, member, Configuration.StylesheetProvider);
            action(columnBuilder);
            base.AddColumn(member, columnBuilder);
            return this;
        }
    }
}
