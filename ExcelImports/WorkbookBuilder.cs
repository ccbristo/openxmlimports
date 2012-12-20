using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelImports.Core;

namespace ExcelImports
{
    public class WorkbookBuilder<TWorkbook>
    {
        private readonly WorkbookConfiguration mConfiguration = new WorkbookConfiguration(typeof(TWorkbook));
        private readonly Dictionary<object, WorksheetBuilder> worksheets = new Dictionary<object, WorksheetBuilder>();
        private INamingConvention WorksheetNamingConvention;

        public WorkbookBuilder()
        {
            WorksheetNamingConvention = new CamelCaseNamingConvention();
        }

        public WorkbookConfiguration Create()
        {
            // TODO [ccb] Do we need a way to differentiate entity types from non-entity types ala NH?
            // Need to differentiate things that will be tables from other properties...
            // Ex:
            // class X{ Something SingleItemProperty; IList<Item> Items { get; set; } }
            // Should pay attention Items and SingleItemProperty.

            var workbookMembers = typeof(TWorkbook).GetMembers(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.IsPropertyOrField() &&
                            m.GetPropertyOrFieldType().ClosesInterface(typeof(IList<>)));

            foreach (var member in workbookMembers)
            {
                var worksheetBuilder = GetTypedWorksheet(member);
                worksheetBuilder.ConfigureColumns();
                this.mConfiguration.AddWorksheet(worksheetBuilder.Configuration);
            }

            return mConfiguration;
        }

        public WorkbookBuilder<TWorkbook> Stylesheet(IStylesheetProvider stylesheetProvider)
        {
            mConfiguration.StylesheetProvider = stylesheetProvider;
            return this;
        }

        public WorkbookBuilder<TWorkbook> Worksheet(string name, Action<WorksheetBuilder> action)
        {
            WorksheetBuilder worksheetConfig = GetNamedWorksheet(name);
            action(worksheetConfig);
            worksheets[name] = worksheetConfig;
            return this;
        }

        private WorksheetBuilder GetNamedWorksheet(string name)
        {
            WorksheetBuilder worksheetConfig;

            if (!worksheets.TryGetValue(name, out worksheetConfig))
            {
                worksheetConfig = new WorksheetBuilder(mConfiguration.StylesheetProvider);
                worksheets[name] = worksheetConfig;
            }

            return worksheetConfig;
        }

        public WorkbookBuilder<TWorkbook> Worksheet<TWorksheet>(
            Expression<Func<TWorkbook, IList<TWorksheet>>> member,
            Action<WorksheetBuilder<TWorksheet>, IStylesheetProvider> action)
        {
            var worksheet = GetTypedWorksheet(member);
            action(worksheet, mConfiguration.StylesheetProvider);
            worksheets[member.GetMemberInfo()] = worksheet;
            return this;
        }

        public WorksheetBuilder<TWorksheet>
            GetWorksheet<TWorksheet>(Expression<Func<TWorkbook, IList<TWorksheet>>> memberExp)
        {
            return GetTypedWorksheet(memberExp);
        }

        private WorksheetBuilder<TWorksheet> GetTypedWorksheet<TWorksheet>(Expression<Func<TWorkbook, IList<TWorksheet>>> memberExp)
        {
            var memberInfo = memberExp.GetMemberInfo();

            WorksheetBuilder worksheetBuilder;
            bool found = worksheets.TryGetValue(memberInfo, out worksheetBuilder);

            if (!found)
            {
                string worksheetName = WorksheetNamingConvention.GetName(memberInfo);
                worksheetBuilder = new WorksheetBuilder<TWorksheet>(worksheetName, memberInfo.Name, mConfiguration.StylesheetProvider);
                worksheets[memberInfo] = worksheetBuilder;
            }

            var castConfig = (WorksheetBuilder<TWorksheet>)worksheetBuilder;
            return castConfig;
        }

        private WorksheetBuilder GetTypedWorksheet(MemberInfo memberInfo)
        {
            var closedCollectionType = memberInfo.GetPropertyOrFieldType().GetClosingInterface(typeof(IList<>));
            var param = Expression.Parameter(memberInfo.DeclaringType, "x");
            var exp = Expression.MakeMemberAccess(param, memberInfo);
            var funcType = Expression.GetFuncType(memberInfo.DeclaringType, closedCollectionType);
            var lambda = Expression.Lambda(funcType, exp, param);

            var methods = this.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic)
                .Where(m => StringComparer.Ordinal.Equals(m.Name, "GetTypedWorksheet") && m.IsGenericMethod)
                .ToList();

            if (methods.Count == 0)
                throw new Exception("Could not find generic overload of GetTypedWorksheet");

            return (WorksheetBuilder)methods[0].MakeGenericMethod(closedCollectionType.GetGenericArguments()[0]).Invoke(this, new[] { lambda });
        }
    }
}
