using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
namespace ExcelImports.Core
{
    public class WorksheetConfiguration
    {
        private readonly List<ColumnConfiguration> mColumns = new List<ColumnConfiguration>();

        public IEnumerable<ColumnConfiguration> Columns
        {
            get { return mColumns.AsReadOnly(); }
        }

        internal WorksheetConfiguration()
            : this(null, null)
        { }

        public WorksheetConfiguration(string sheetName, string memberName)
        {
            this.SheetName = sheetName;
            this.MemberName = memberName;
        }

        public bool ExportOnly { get; set; }
        public string SheetName { get; set; }
        public string MemberName { get; set; }

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

        internal ICollection GetMember(object workbookSource)
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

            var memberInfo = memberInfos[0];
            object value = memberInfo.GetPropertyOrFieldValue(workbookSource);

            if (!(value is ICollection))
                // TODO [ccb] Improve this message
                throw new InvalidOperationException("Worksheet source is not an ICollection");

            return (ICollection)value;
        }
    }
}
