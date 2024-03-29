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

namespace OpenXmlImports.Core
{
    public class Importer
    {
        public TWorkbook Import<TWorkbook>(WorkbookConfiguration<TWorkbook> workbookConfiguration, Stream input)
        {
            var result = (TWorkbook)Create(workbookConfiguration.BoundType);
            var document = CreateDocument(input, workbookConfiguration.ErrorPolicy);
            var sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
            var sheets = document.WorkbookPart.Workbook.Sheets;

            // TODO [ccb] Add tests for import only sheets
            foreach (var worksheetConfig in workbookConfiguration)
            {
                var sheet = worksheetConfig.GetWorksheet(sheets, workbookConfiguration.ErrorPolicy);
                var worksheetMember = workbookConfiguration.GetMemberInfoFor(worksheetConfig);

                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var headerRow = sheetData.Elements<Row>().First();

                var columnMap = MapColumns(worksheetConfig, workbookConfiguration.ErrorPolicy,
                    headerRow, sharedStringTable);
                var dataRows = sheetData.Elements<Row>().Skip(1).Where(HasCellValues).ToList();

                var memberType = worksheetMember.GetMemberType();

                if (memberType.ClosesInterface(typeof(IList<>)))
                {
                    var list = (IList)Create(memberType);
                    worksheetMember.SetPropertyOrFieldValue(result, list);
                    AddWorksheetItems(dataRows, list, worksheetConfig,
                        columnMap, sharedStringTable, workbookConfiguration.ErrorPolicy);
                }
                else if (string.IsNullOrEmpty(worksheetConfig.MemberName) && dataRows.Count == 1)
                {
                    ImportItem(result, worksheetConfig, columnMap, sharedStringTable,
                        workbookConfiguration.ErrorPolicy, dataRows[0]);
                }
                else
                {
                    var target = Create(memberType);
                    ImportItem(target, worksheetConfig, columnMap, sharedStringTable,
                        workbookConfiguration.ErrorPolicy, dataRows[0]);
                    worksheetMember.SetPropertyOrFieldValue(result, target);
                }
            }

            workbookConfiguration.ErrorPolicy.OnImportComplete();
            return result;
        }

        private static bool HasCellValues(Row row)
        {
            return row.Descendants<CellValue>().Any();
        }

        private IDictionary<ColumnConfiguration, ColumnReference> MapColumns(WorksheetConfiguration worksheetConfig,
            IErrorPolicy errorPolicy, Row headerRow, SharedStringTable sharedStrings)
        {
            var map = new Dictionary<ColumnConfiguration, ColumnReference>();

            // only pay attention to columns configured in.
            // ignore any extra columns in the sheet and columns explicitly ignored.
            foreach (var column in worksheetConfig.NonIgnoredColumns)
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
            foreach (var row in dataRows)
            {
                var item = Create(worksheetConfig.BoundType);
                ImportItem(item, worksheetConfig, columnMap, sharedStrings, errorPolicy, row);
                list.Add(item);
            }
        }

        private static void ImportItem(
            object target, WorksheetConfiguration worksheetConfig,
            IDictionary<ColumnConfiguration, ColumnReference> columnMap,
            SharedStringTable sharedStrings, IErrorPolicy errorPolicy, Row row)
        {
            // only pay attention to columns configured in.
            // ignore any extra columns in the sheet and any explicitly ignored columns.
            foreach (var column in worksheetConfig.NonIgnoredColumns)
            {
                var colRef = columnMap[column];
                var cell = row.Elements<Cell>().SingleOrDefault(c => c.CellReference.Column() == colRef);

                bool cellHasValue = cell != null && cell.CellValue != null;

                if (!cellHasValue && column.Required)
                {
                    errorPolicy.OnRequiredColumnViolation(worksheetConfig.SheetName, column.Name,
                        colRef, row.RowIndex);
                }
                else if (cellHasValue)
                {
                    CellValues? dataType = cell.DataType?.Value;

                    string text = cell.CellValue?.Text;
                    if (dataType == CellValues.SharedString)
                        text = sharedStrings.GetText(cell.CellValue);

                    if (text.StartsWith("'"))
                        text = text.Substring(1);

                    object value = column.Type.NullSafeGet(text);

                    if (column.Member.GetMemberType().Is<string>() &&
                        ((string)value ?? string.Empty).Length > column.MaxLength)
                        errorPolicy.OnMaxLengthExceeded(colRef, row.RowIndex, column.MaxLength, column.Name);

                    column.SetValue(target, value);
                }
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
                throw new InvalidOperationException(
                    $"The requested type {type.FullName} does not have a zero-argument constructor.");

            return ctor.Invoke(new object[] { });
        }
    }
}
