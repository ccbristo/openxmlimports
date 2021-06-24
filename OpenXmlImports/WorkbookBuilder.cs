using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OpenXmlImports.Core;

namespace OpenXmlImports
{
    public class WorkbookBuilder<TWorkbook, TStylesheet>
        where TStylesheet : IStylesheetProvider
    {
        private readonly WorkbookConfiguration<TWorkbook> mConfiguration;
        private readonly Dictionary<object, WorksheetBuilder> worksheets = new Dictionary<object, WorksheetBuilder>();
        private readonly INamingConvention WorksheetNamingConvention;

        public WorkbookBuilder(TStylesheet stylesheet)
        {
            WorksheetNamingConvention = new CamelCaseNamingConvention();
            mConfiguration = new WorkbookConfiguration<TWorkbook>(stylesheet);
        }

        public WorkbookConfiguration<TWorkbook> Create()
        {
            var bindingFlags = BindingFlags.Public | BindingFlags.Instance;

            var rootProperties = typeof(TWorkbook).GetMembers(bindingFlags)
                .Where(m => m.IsPropertyOrField() && m.GetMemberType().IsTerminal())
                .ToList();

            if (rootProperties.Any())
            {
                var rootPropertiesBuilder = GetRootWorksheetBuilder();
                rootPropertiesBuilder.ConfigureColumns();
                mConfiguration.AddWorksheet(rootPropertiesBuilder.Configuration);
            }

            var nonRootProperties = typeof(TWorkbook).GetMembers(bindingFlags)
                .Except(rootProperties)
                .Where(m => m.IsPropertyOrField());

            foreach (var member in nonRootProperties)
            {
                WorksheetBuilder builder;
                worksheets.TryGetValue(member, out builder);

                if (builder == null)
                {
                    string name = this.WorksheetNamingConvention.GetName(member);
                    builder = WorksheetBuilder.Create(name, member, mConfiguration.StylesheetProvider);
                }

                builder.ConfigureColumns();
                this.mConfiguration.AddWorksheet(builder.Configuration);
            }

            return mConfiguration;
        }

        public WorkbookBuilder<TWorkbook, TStylesheet> Worksheet(string name, Action<WorksheetBuilder> configure)
        {
            var worksheetBuilder = GetNamedWorksheet(name);
            configure(worksheetBuilder);
            worksheets[name] = worksheetBuilder;
            return this;
        }

        private WorksheetBuilder GetNamedWorksheet(string name)
        {
            if (!worksheets.TryGetValue(name, out var worksheetConfig))
            {
                worksheetConfig = new WorksheetBuilder(mConfiguration.StylesheetProvider);
                worksheets[name] = worksheetConfig;
            }

            return worksheetConfig;
        }

        public WorkbookBuilder<TWorkbook, TStylesheet> Singleton<TWorksheet>(
            Expression<Func<TWorkbook, TWorksheet>> member,
            Action<WorksheetBuilder<TWorksheet>, TStylesheet> configure)
        {
            var memberInfo = member.GetMemberInfo();
            string name = this.WorksheetNamingConvention.GetName(memberInfo);
            var worksheet = new WorksheetBuilder<TWorksheet>(name, memberInfo.Name, mConfiguration.StylesheetProvider);

            configure(worksheet, (TStylesheet)mConfiguration.StylesheetProvider);
            worksheets[memberInfo] = worksheet;
            return this;
        }

        public WorkbookBuilder<TWorkbook, TStylesheet> List<TWorksheet>(
            Expression<Func<TWorkbook, IList<TWorksheet>>> list,
            Action<WorksheetBuilder<TWorksheet>, TStylesheet> configure)
        {
            var member = list.GetMemberInfo();
            string name = WorksheetNamingConvention.GetName(member);
            var worksheetBuilder = new WorksheetBuilder<TWorksheet>(name, member.Name, mConfiguration.StylesheetProvider);
            configure(worksheetBuilder, (TStylesheet)mConfiguration.StylesheetProvider);
            worksheets[member] = worksheetBuilder;
            return this;
        }

        public WorkbookBuilder<TWorkbook, TStylesheet> RootProperties(Action<WorksheetBuilder<TWorkbook>, TStylesheet> configure)
        {
            var builder = GetRootWorksheetBuilder();
            configure(builder, (TStylesheet)mConfiguration.StylesheetProvider);
            return this;
        }

        private WorksheetBuilder<TWorkbook> GetRootWorksheetBuilder()
        {
            if (worksheets.TryGetValue(typeof(TWorkbook), out var worksheetBuilder))
                return (WorksheetBuilder<TWorkbook>)worksheetBuilder;

            var result = new WorksheetBuilder<TWorkbook>("Details", mConfiguration.StylesheetProvider);
            worksheets[typeof(TWorkbook)] = result;

            return result;
        }
    }
}
