using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelImports.Core
{
    public class WorkbookConfiguration : IEnumerable<WorksheetConfiguration>
    {
        public string Name { get; set; }
        public Type BoundType { get; private set; }
        public IErrorPolicy ErrorPolicy { get; set; }
        public IStylesheetProvider StylesheetProvider { get; set; }
        private readonly List<WorksheetConfiguration> Worksheets = new List<WorksheetConfiguration>();

        public WorkbookConfiguration(Type boundType)
        {
            BoundType = boundType;
            StylesheetProvider = new DefaultStylesheetProvider();
            ErrorPolicy = new ImmediateExceptionErrorPolicy();
        }

        public void AddWorksheet(WorksheetConfiguration worksheet)
        {
            this.Worksheets.Add(worksheet);
        }

        public WorksheetConfiguration GetWorksheet(string name)
        {
            return Worksheets.SingleOrDefault(c => StringComparer.OrdinalIgnoreCase.Equals(name, c.SheetName));
        }

        internal MemberInfo GetMemberInfoFor(WorksheetConfiguration worksheet)
        {
            var memberInfos = BoundType.GetMember(worksheet.MemberName, BindingFlags.Public | BindingFlags.Instance);

            if (memberInfos.Length == 0)
                throw new MissingMemberException(string.Format("Could not find a public instance member named {0} on {1}",
                    worksheet.MemberName, BoundType.Name));
            else if (memberInfos.Length > 1)
                // TODO [ccb] Verify this is the type of exception to throw and add a meaningful message.
                throw new AmbiguousMatchException();

            return memberInfos[0];
        }

        internal IList GetListFor(WorksheetConfiguration worksheetConfig, object workbookSource)
        {
            var memberInfo = GetMemberInfoFor(worksheetConfig);
            object value = memberInfo.GetPropertyOrFieldValue(workbookSource);

            if (!(value is IList))
                // TODO [ccb] Improve this message
                throw new InvalidOperationException("Worksheet source is not an IList");

            return (IList)value;
        }

        public void Export(object workbookSource, Stream output)
        {
            ExcelExporter exporter = new ExcelExporter();
            exporter.Export(this, workbookSource, output);
        }

        public object Import(Stream input)
        {
            ExcelImporter importer = new ExcelImporter();
            return importer.Import(this, input);
        }

        public IEnumerator<WorksheetConfiguration> GetEnumerator()
        {
            return Worksheets.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return Worksheets.GetEnumerator();
        }
    }
}
