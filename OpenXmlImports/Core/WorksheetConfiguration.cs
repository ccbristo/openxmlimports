using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlImports.Types;

namespace OpenXmlImports.Core
{
    public class WorksheetConfiguration
    {
        private readonly Dictionary<MemberInfo, ColumnConfiguration> mColumns = new Dictionary<MemberInfo, ColumnConfiguration>();
        internal IStylesheetProvider StylesheetProvider { get; set; }

        public IEnumerable<ColumnConfiguration> Columns => mColumns.Values;

        internal IEnumerable<ColumnConfiguration> NonIgnoredColumns => mColumns.Values.Where(c => !c.Ignore);

        internal WorksheetConfiguration(IStylesheetProvider stylesheetProvider)
            : this(null, null, null, stylesheetProvider)
        { }

        internal WorksheetConfiguration(Type boundType, IStylesheetProvider stylesheetProvider)
            : this(boundType, null, null, stylesheetProvider)
        { }

        public WorksheetConfiguration(Type boundType, string sheetName, string memberName,
            IStylesheetProvider stylesheetProvider)
        {
            this.SheetName = sheetName;
            this.MemberName = memberName;
            this.BoundType = boundType;
            this.StylesheetProvider = stylesheetProvider;
        }

        public Type BoundType { get; set; }
        public string MemberName { get; set; }
        public string SheetName { get; set; }
        public bool ExportOnly { get; set; }

        public ColumnConfiguration GetColumn(MemberInfo member)
        {
            return Columns.SingleOrDefault(c => c.Member == member);
        }

        public ColumnConfiguration AddColumn(string columnName, MemberInfo member)
        {
            var column = new ColumnConfiguration
            {
                Name = columnName,
                Member = member,
                Type = TypeFactory.GetType(member.GetMemberType())
            };

            return AddColumn(column);
        }

        public ColumnConfiguration AddColumn(ColumnConfiguration column)
        {
            mColumns.Add(column.Member, column);
            return column;
        }

        internal Sheet GetWorksheet(Sheets sheets, IErrorPolicy errorPolicy)
        {
            var matchingSheets = sheets.Elements<Sheet>()
                .Where(sheet => string.Equals(this.SheetName, sheet.Name))
                .ToList();

            if (matchingSheets.Count != 1)
                errorPolicy.OnMissingWorksheet(this.SheetName);

            return matchingSheets[0];
        }

        public IEnumerator<ColumnConfiguration> GetEnumerator()
        {
            return mColumns.Values.GetEnumerator();
        }
    }
}
