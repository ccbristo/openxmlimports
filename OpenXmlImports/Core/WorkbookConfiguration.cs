using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace OpenXmlImports.Core
{
    public class WorkbookConfiguration
    {
        public string Name { get; set; }
        public Type BoundType { get; private set; }
        public IErrorPolicy ErrorPolicy { get; set; }
        public IStylesheetProvider StylesheetProvider { get; private set; }
        private readonly Dictionary<object, WorksheetConfiguration> mWorksheets = new Dictionary<object, WorksheetConfiguration>();

        public WorkbookConfiguration(Type boundType, IStylesheetProvider stylesheetProvider)
        {
            BoundType = boundType;
            StylesheetProvider = stylesheetProvider;
            ErrorPolicy = new ImmediateExceptionErrorPolicy();
        }

        public void AddWorksheet(WorksheetConfiguration worksheet)
        {
            mWorksheets[worksheet.SheetName] = worksheet;
        }

        public WorksheetConfiguration GetWorksheet(string name)
        {
            WorksheetConfiguration config;
            mWorksheets.TryGetValue(name, out config);
            return config;
        }

        public WorksheetConfiguration GetWorksheet(MemberInfo member)
        {
            WorksheetConfiguration config;
            mWorksheets.TryGetValue(member, out config);
            return config;
        }

        internal MemberInfo GetMemberInfoFor(WorksheetConfiguration worksheet)
        {
            if (string.IsNullOrEmpty(worksheet.MemberName))
                return BoundType;

            var memberInfos = BoundType.GetMember(worksheet.MemberName, BindingFlags.Public | BindingFlags.Instance);

            if (memberInfos.Length == 0)
                throw new MissingMemberException(string.Format("Could not find a public instance member named {0} on {1}",
                    worksheet.MemberName, BoundType.Name));
            else if (memberInfos.Length > 1)
                // TODO [ccb] Verify this is the type of exception to throw and add a meaningful message.
                throw new AmbiguousMatchException();

            return memberInfos[0];
        }

        internal object GetMemberFor(WorksheetConfiguration worksheetConfig, object workbookSource)
        {
            if (string.IsNullOrEmpty(worksheetConfig.MemberName))
                return workbookSource;

            var memberInfo = GetMemberInfoFor(worksheetConfig);
            object value = memberInfo.GetPropertyOrFieldValue(workbookSource);
            return value;
        }

        public void Export(object workbookSource, Stream output)
        {
            Exporter exporter = new Exporter();
            exporter.Export(this, workbookSource, output);
        }

        public object Import(Stream input)
        {
            Importer importer = new Importer();
            return importer.Import(this, input);
        }

        public IEnumerable<WorksheetConfiguration> Worksheets
        {
            get { return mWorksheets.Values; }
        }

        public IEnumerator<WorksheetConfiguration> GetEnumerator()
        {
            return mWorksheets.Values.GetEnumerator();
        }
    }
}
