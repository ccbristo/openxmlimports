﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace ExcelImports.Core
{
    public class ExcelImporter
    {
        public object Import(WorkbookConfiguration workbookConfiguration, Stream input)
        {
            object result = Create(workbookConfiguration.BoundType);
            var document = CreateDocument(input, workbookConfiguration.ErrorPolicy);

            var sheets = document.WorkbookPart.Workbook.Sheets;

            // TODO [ccb] Add tests for import only sheets
            foreach (var worksheetConfig in workbookConfiguration)
            {
                var sheet = worksheetConfig.GetWorksheet(sheets, workbookConfiguration.ErrorPolicy);
                var worksheetMember = workbookConfiguration.GetMemberInfoFor(worksheetConfig);
                var list = (IList)Create(worksheetMember.GetPropertyOrFieldType());
                worksheetMember.SetPropertyOrFieldValue(result, list);

                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);

                var sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var headerRow = sheetData.Elements<Row>().First();
                var dataRows = sheetData.Elements<Row>().Skip(1).ToList();

                var columnMap = MapColumns(worksheetConfig, workbookConfiguration.ErrorPolicy,
                    headerRow, sharedStringTable);
                AddWorksheetItems(dataRows, list, worksheetConfig,
                    columnMap, sharedStringTable, workbookConfiguration.ErrorPolicy);

            }

            workbookConfiguration.ErrorPolicy.OnImportComplete();
            return result;
        }

        private IDictionary<ColumnConfiguration, ColumnReference> MapColumns(WorksheetConfiguration worksheetConfig,
            IErrorPolicy errorPolicy, Row headerRow, SharedStringTable sharedStrings)
        {
            var map = new Dictionary<ColumnConfiguration, ColumnReference>();

            // only pay attention to columns configured in.
            // ignore any extra columns in the sheet.
            foreach (var column in worksheetConfig.Columns)
            {
                var cells = headerRow.Elements<Cell>().Where(c => c.GetCellText(sharedStrings) == column.Name)
                    .ToList();

                if (cells.Count == 0)
                    errorPolicy.OnMissingColumn(column.Name);
                else if (cells.Count > 1)
                    errorPolicy.OnDuplicatedColumn(column.Name);

                map.Add(column, cells[0].CellReference.Column());
            }

            return map;
        }

        private void AddWorksheetItems(IEnumerable<Row> dataRows,
            IList list,
            WorksheetConfiguration worksheetConfig,
            IDictionary<ColumnConfiguration, ColumnReference> columnMap,
            SharedStringTable sharedStrings,
            IErrorPolicy errorPolicy)
        {
            int rowIndex = 2; // data starts on row 2

            foreach (var row in dataRows)
            {
                var item = Create(worksheetConfig.BoundType);

                // only pay attention to columns configured in.
                // ignore any extra columns in the sheet.
                foreach (var column in worksheetConfig.Columns)
                {
                    var colRef = columnMap[column];
                    var cell = row.Elements<Cell>().SingleOrDefault(c => c.CellReference.Column() == colRef);

                    bool cellHasValue = cell != null && cell.CellValue != null;

                    if (!cellHasValue && !column.AllowNull)
                    {
                        errorPolicy.OnNullableColumnViolation(worksheetConfig.SheetName, column.Name,
                            colRef, rowIndex);
                        continue;
                    }
                    else if (cellHasValue)
                    {
                        string text = cell.GetCellText(sharedStrings);
                        column.SetValue(item, text);
                    }
                }

                rowIndex++;
                list.Add(item);
            }
        }

        private static SpreadsheetDocument CreateDocument(Stream input, IErrorPolicy errorPolicy)
        {
            var document = SpreadsheetDocument.Open(input, false);

            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Office2007);
            var errorInfo = validator.Validate(document);

            if (errorInfo.Any())
            {
                errorPolicy.OnInvalidFile(errorInfo);
                errorPolicy.OnImportComplete(); // we can't continue if the file is invalid.

                // TODO [ccb] the below seems fishy, but it also seems that
                // a correctly written error policy will either throw an exception
                // immediately upon being notified of an invalidty or
                // would do so during OnImportComplete, which means this statement
                // should never be hit, in any case, the below ensures processing ends here in such a case.
                throw new InvalidImportFileException("The document is invalid.");
            }

            return document;
        }

        private static object Create(Type type)
        {
            var ctor = type.GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                Type.DefaultBinder, Type.EmptyTypes, null);

            if (ctor == null)
                throw new InvalidOperationException(string.Format(
                    "The requested type {0} does not have a zero-argument constructor.",
                    type.FullName));

            return ctor.Invoke(new object[] { });
        }
    }
}
