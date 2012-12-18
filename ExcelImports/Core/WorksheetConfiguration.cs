using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImports.Core
{
    public class WorksheetConfiguration : IEnumerable<ColumnConfiguration>
    {
        private readonly List<ColumnConfiguration> mColumns = new List<ColumnConfiguration>();

        public IEnumerable<ColumnConfiguration> Columns
        {
            get { return mColumns.AsReadOnly(); }
        }

        internal WorksheetConfiguration(Type boundType)
            : this(boundType, null, null)
        { }

        public WorksheetConfiguration(Type boundType, string sheetName, string memberName)
        {
            this.SheetName = sheetName;
            this.MemberName = memberName;
            this.BoundType = boundType;
        }

        public Type BoundType { get; set; }
        public string MemberName { get; set; }
        public string SheetName { get; set; }
        public bool ExportOnly { get; set; }

        public ColumnConfiguration AddColumn(string columnName, string memberName)
        {
            var column = new ColumnConfiguration
            {
                Name = columnName,
                MemberName = memberName
            };

            mColumns.Add(column);

            return column;
        }

        internal Sheet GetWorksheet(Sheets sheets)
        {
            var matchingSheets = sheets.Elements<Sheet>()
                .Where(sheet => string.Equals(this.SheetName, sheet.Name))
                .ToList();

            if (matchingSheets.Count != 1) // use error policy
                throw new MissingWorksheetException("No sheet named \"{0}\" was found in the workbook.",
                    this.SheetName);

            return matchingSheets[0];
        }

        internal MemberInfo GetMemberInfo(object workbookSource)
        {
            if (workbookSource == null)
                throw new ArgumentNullException("workbookSource");

            var memberInfos = workbookSource.GetType()
                .GetMember(this.MemberName, BindingFlags.Public | BindingFlags.Instance);

            if (memberInfos.Length == 0)
                throw new MissingMemberException(string.Format("Could not find a public instance member named {0} on {1}",
                    this.MemberName, workbookSource.GetType().Name));
            else if (memberInfos.Length > 1)
                // TODO [ccb] Verify this is the type of exception to throw and add a meaningful message.
                throw new AmbiguousMatchException();

            return memberInfos[0];
        }

        internal IList GetMember(object workbookSource)
        {
            var memberInfo = GetMemberInfo(workbookSource);
            object value = memberInfo.GetPropertyOrFieldValue(workbookSource);

            if (!(value is IList))
                // TODO [ccb] Improve this message
                throw new InvalidOperationException("Worksheet source is not an IList");

            return (IList)value;
        }

        public IEnumerator<ColumnConfiguration> GetEnumerator()
        {
            return mColumns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
