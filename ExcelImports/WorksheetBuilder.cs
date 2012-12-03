using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using ExcelImports.Core;

namespace ExcelImports
{
    public class WorksheetBuilder
    {
        public WorksheetBuilder()
            : this(new WorksheetConfiguration())
        {

        }

        protected WorksheetBuilder(WorksheetConfiguration config)
        {
            this.Configuration = config;
        }

        public WorksheetConfiguration Configuration { get; private set; }

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
            this.Configuration.Name = name;
            return this;
        }
    }

    public class WorksheetBuilder<T>
        : WorksheetBuilder
    {
        public WorksheetBuilder()
            : base(new WorksheetConfiguration<T>())
        { }

        readonly Dictionary<MemberInfo, ColumnConfiguration> ColumnConfigurations = new Dictionary<MemberInfo, ColumnConfiguration>();

        public new WorksheetBuilder<T> Named(string name)
        {
            return (WorksheetBuilder<T>)base.Named(name);
        }

        public WorksheetBuilder<T> Column<TColumn>(Expression<Func<T, TColumn>> columnExp,
            Action<ColumnConfiguration<TColumn>> action)
        {
            throw new NotImplementedException();

            //bool isMultiColumn = IsMultiColumn(columnExp);

            //var memberInfo = columnExp.GetMemberInfo();
            //ColumnConfiguration columnConfig;
            //ColumnConfigurations.TryGetValue(memberInfo, out columnConfig);

            //if (columnConfig == null)
            //{
            //    columnConfig = new ColumnConfiguration<TColumn>();
            //    ColumnConfigurations[memberInfo] = columnConfig;
            //}

            //return this;
        }

        //private bool IsMultiColumn(Expression exp)
        //{
        //    var lamdaExp = ((LambdaExpression)exp);
        //    return lamdaExp.Body.NodeType == ExpressionType.New;
        //}
    }
}
