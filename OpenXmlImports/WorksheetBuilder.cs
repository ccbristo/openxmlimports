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

        internal WorksheetBuilder(IStylesheetProvider stylesheetProvider)
            : this(new WorksheetConfiguration(stylesheetProvider))
        { }

        internal WorksheetBuilder(Type boundType, string sheetName, string memberName, IStylesheetProvider stylesheetProvider)
            : this(new WorksheetConfiguration(boundType, sheetName, memberName, stylesheetProvider))
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

        internal void ConfigureColumns()
        {
            var columnMembers = Configuration.BoundType.GetMembers(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.IsPropertyOrField() && m.GetMemberType().IsTerminal())
                .ToList();

            foreach (var member in columnMembers)
            {
                var columnConfig = Configuration.GetColumn(member);

                if (columnConfig == null)
                {
                    string name = ColumnNamingConvention.GetName(member);
                    var columnBuilder = ColumnBuilder.For(member, name, Configuration.StylesheetProvider);
                    Configuration.AddColumn(columnBuilder.Configuration);
                }
            }
        }

        internal static WorksheetBuilder Create(string name, MemberInfo member, IStylesheetProvider stylesheetProvider)
        {
            Type sheetType = member.GetMemberType();

            if (sheetType.ClosesInterface(typeof(IList<>)))
                sheetType = sheetType.GetClosingInterface(typeof(IList<>)).GetGenericArguments().Single();

            var worksheetBuilder = new WorksheetBuilder(sheetType, name, member.Name, stylesheetProvider);
            return worksheetBuilder;
        }
    }

    public class WorksheetBuilder<T>
        : WorksheetBuilder
    {
        public WorksheetBuilder(string name, string memberName, IStylesheetProvider stylesheetProvider)
            : base(typeof(T), name, memberName, stylesheetProvider)
        { }

        internal WorksheetBuilder(string name, IStylesheetProvider stylesheet)
            : this(name, null, stylesheet)
        { }

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
            Configuration.AddColumn(columnBuilder.Configuration);
            return this;
        }
    }
}
