using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using OpenXmlImports.Core;

namespace OpenXmlImports
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
            var bindingFlags = BindingFlags.Public | BindingFlags.Instance;

            var rootProperties = typeof(TWorkbook).GetMembers(bindingFlags)
                .Where(m => m.IsPropertyOrField() && m.GetMemberType().IsTerminal());

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

        public WorkbookBuilder<TWorkbook> Singleton<TWorksheet>(
            Expression<Func<TWorkbook, TWorksheet>> member,
            Action<WorksheetBuilder<TWorksheet>, IStylesheetProvider> configure)
        {
            var memberInfo = member.GetMemberInfo();
            string name = this.WorksheetNamingConvention.GetName(memberInfo);
            var worksheet = new WorksheetBuilder<TWorksheet>(name, memberInfo.Name, mConfiguration.StylesheetProvider);

            configure(worksheet, mConfiguration.StylesheetProvider);
            worksheets[memberInfo] = worksheet;
            return this;
        }

        public WorkbookBuilder<TWorkbook> List<TWorksheet>(
            Expression<Func<TWorkbook, IList<TWorksheet>>> list,
            Action<WorksheetBuilder<TWorksheet>, IStylesheetProvider> configure)
        {
            var member = list.GetMemberInfo();
            string name = WorksheetNamingConvention.GetName(member);
            var worksheetBuilder = new WorksheetBuilder<TWorksheet>(name, mConfiguration.StylesheetProvider);
            configure(worksheetBuilder, mConfiguration.StylesheetProvider);
            worksheets[member] = worksheetBuilder;
            return this;
        }

        public WorkbookBuilder<TWorkbook> RootProperties(Action<WorksheetBuilder<TWorkbook>, IStylesheetProvider> configure)
        {
            var builder = GetRootWorksheetBuilder();
            configure(builder, mConfiguration.StylesheetProvider);
            return this;
        }

        private WorksheetBuilder<TWorkbook> GetRootWorksheetBuilder()
        {
            WorksheetBuilder worksheetBuilder;

            if (worksheets.TryGetValue(typeof(TWorkbook), out worksheetBuilder))
                return (WorksheetBuilder<TWorkbook>)worksheetBuilder;

            var result = new WorksheetBuilder<TWorkbook>("Details", mConfiguration.StylesheetProvider);
            worksheets[typeof(TWorkbook)] = result;

            return result;
        }
    }
}
