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
        private readonly WorkbookConfiguration<TWorkbook> mConfiguration = new WorkbookConfiguration<TWorkbook>();
        private readonly Dictionary<object, WorksheetBuilder> worksheets = new Dictionary<object, WorksheetBuilder>();
        private Func<TWorkbook, string> namer;

        public WorkbookConfiguration<TWorkbook> Create()
        {
            // TODO [ccb] Do we need a way to differentiate entity types from non-entity types ala NH?
            // Need to differentiate things that will be tables from other properties...
            // Ex:
            // class X{ string S; ICollection<Item> Items { get; set; } }
            // Should only pay attention to Items.

            var workbookMembers = typeof(TWorkbook).GetMembers(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.IsPropertyOrField() &&
                            m.GetPropertyOrFieldType().ClosesInterface(typeof(ICollection<>)));

            foreach (var member in workbookMembers)
            {
                var worksheet = GetTypedWorksheet(member);
                this.mConfiguration.AddWorksheet(worksheet.Configuration);
            }

            return mConfiguration;
        }

        public WorkbookBuilder<TWorkbook> Named(string name)
        {
            return Named(t => name);
        }

        public WorkbookBuilder<TWorkbook> Named(Func<TWorkbook, string> namer)
        {
            this.namer = namer;
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
                worksheetConfig = new WorksheetBuilder();
                worksheets[name] = worksheetConfig;
            }
            return worksheetConfig;
        }

        public WorkbookBuilder<TWorkbook> Worksheet<TWorksheet>(
            Expression<Func<TWorkbook, ICollection<TWorksheet>>> member,
            Action<WorksheetBuilder<TWorksheet>> action)
        {
            var worksheet = GetTypedWorksheet(member);
            action(worksheet);
            worksheets[member.GetMemberInfo()] = worksheet;
            return this;
        }

        public WorksheetBuilder<TWorksheet>
            GetWorksheet<TWorksheet>(Expression<Func<TWorkbook, ICollection<TWorksheet>>> memberExp)
        {
            return GetTypedWorksheet(memberExp);
        }

        private WorksheetBuilder<TWorksheet> GetTypedWorksheet<TWorksheet>(Expression<Func<TWorkbook, ICollection<TWorksheet>>> memberExp)
        {
            var memberInfo = memberExp.GetMemberInfo();

            WorksheetBuilder worksheetConfig;
            bool found = worksheets.TryGetValue(memberInfo, out worksheetConfig);

            if (!found)
            {
                worksheetConfig = new WorksheetBuilder<TWorksheet>();
                worksheets[memberInfo] = worksheetConfig;
            }

            var castConfig = (WorksheetBuilder<TWorksheet>)worksheetConfig;
            return castConfig;
        }

        private WorksheetBuilder GetTypedWorksheet(MemberInfo memberInfo)
        {
            var closedCollectionType = memberInfo.GetPropertyOrFieldType().GetClosingInterface(typeof(ICollection<>));
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
