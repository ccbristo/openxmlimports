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
        internal IStylesheetProvider StylesheetProvider { get; set; }

        public IEnumerable<ColumnConfiguration> Columns
        {
            get { return mColumns.AsReadOnly(); }
        }

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

        public ColumnConfiguration AddColumn(string columnName, MemberInfo member)
        {
            var column = new ColumnConfiguration
            {
                Name = columnName,
                Member = member
            };

            return AddColumn(column);
        }

        public ColumnConfiguration AddColumn(ColumnConfiguration column)
        {
            mColumns.Add(column);
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
            return mColumns.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
