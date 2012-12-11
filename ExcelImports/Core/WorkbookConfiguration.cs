﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelImports.Core
{
    public class WorkbookConfiguration : IEnumerable<WorksheetConfiguration>
    {
        public string Name { get; set; }
        public IStylesheetProvider StylesheetProvider { get; set; }
        private readonly List<WorksheetConfiguration> Worksheets = new List<WorksheetConfiguration>();

        public WorkbookConfiguration()
        {
            StylesheetProvider = new DefaultStylesheetProvider();
        }

        public void AddWorksheet(WorksheetConfiguration worksheet)
        {
            this.Worksheets.Add(worksheet);
        }

        public WorksheetConfiguration GetWorksheet(string name)
        {
            return Worksheets.SingleOrDefault(c => StringComparer.OrdinalIgnoreCase.Equals(name, c.SheetName));
        }

        public void Export(object workbookSource, Stream output)
        {
            ExcelExporter exporter = new ExcelExporter();
            exporter.Export(this, workbookSource, output);
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

    public sealed class WorkbookConfiguration<TWorkbook> : WorkbookConfiguration
    {

    }
}
