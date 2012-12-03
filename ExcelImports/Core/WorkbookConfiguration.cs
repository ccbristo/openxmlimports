using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelImports.Core
{
    public sealed class WorkbookConfiguration<TWorkbook>
    {
        public string Name { get; set; }
        public INamingConvention WorksheetNamingConvention { get; set; }

        private readonly List<WorksheetConfiguration> Worksheets = new List<WorksheetConfiguration>();

        public WorkbookConfiguration()
        {
            this.WorksheetNamingConvention = new CamelCaseNamingConvention();
        }

        public void AddWorksheet(WorksheetConfiguration worksheet)
        {
            this.Worksheets.Add(worksheet);
        }

        public WorksheetConfiguration GetWorksheet(string name)
        {
            return Worksheets.SingleOrDefault(c => StringComparer.OrdinalIgnoreCase.Equals(name, c.Name));
        }

        public WorksheetConfiguration<TWorksheet> GetWorksheet<TWorksheet>(Expression<Func<TWorkbook, IEnumerable<TWorksheet>>> memberExp)
        {
            string name = WorksheetNamingConvention.GetName(memberExp.GetMemberInfo());
            return (WorksheetConfiguration<TWorksheet>)GetWorksheet(name);
        }
    }
}
