using System.Collections;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlImports.Core
{
    internal class Exporter
    {
        public void Export<TWorkbook>(WorkbookConfiguration<TWorkbook> workbookConfig, TWorkbook workbookSource, Stream output)
        {
            var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            var sharedStringTablePart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = (Stylesheet)workbookConfig.StylesheetProvider.Stylesheet.CloneNode(true);

            var sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            uint sheetId = 1;

            foreach (var worksheetConfig in workbookConfig)
            {
                SheetData sheetData;
                var sheet = AddSheet(workbookPart, sheetId, worksheetConfig.SheetName, out sheetData);

                var headerRow = CreateHeaderRow(worksheetConfig);
                sheetData.Append(headerRow);

                object source = workbookConfig.GetMemberFor(worksheetConfig, workbookSource);

                if (source is IList)
                    ExportList((IList)source, worksheetConfig, sheetData,
                        workbookConfig.StylesheetProvider, sharedStringTablePart.SharedStringTable);
                else
                    ExportItem(source, worksheetConfig, sheetData, workbookConfig.StylesheetProvider,
                        sharedStringTablePart.SharedStringTable);

                sheets.Append(sheet);
                sheetId++;
            }

            workbookPart.Workbook.Save();
            document.Close();
        }

        private Row CreateHeaderRow(WorksheetConfiguration worksheetConfig)
        {
            var row = new Row();

            ColumnReference colRef = "A";
            foreach (var column in worksheetConfig.NonIgnoredColumns)
            {
                Cell cell = new Cell()
                {
                    CellReference = colRef,
                    CellValue = new CellValue(column.Name),
                    DataType = new EnumValue<CellValues>(CellValues.String),
                };

                row.Append(cell);
                colRef++;
            }

            return row;
        }

        private void ExportList(IList list, WorksheetConfiguration worksheetConfig, SheetData sheetData,
            IStylesheetProvider stylesheet, SharedStringTable sharedStrings)
        {
            foreach (object item in list)
            {
                ExportItem(item, worksheetConfig, sheetData, stylesheet, sharedStrings);
            }
        }

        private void ExportItem(object source, WorksheetConfiguration worksheetConfig,
            SheetData sheetData, IStylesheetProvider stylesheet, SharedStringTable sharedStrings)
        {
            Row row = new Row();
            ColumnReference colRef = "A";

            foreach (var column in worksheetConfig.NonIgnoredColumns)
            {
                object value = column.GetValue(source);

                if (value == null && column.Required)
                    throw new RequiredColumnViolationException("Cannot export null value into non-nullable column \"{0}\" on sheet \"{1}\".",
                        column.Name, worksheetConfig.SheetName);

                if (column.Member.GetMemberType().Is<string>() &&
                    value != null && ((string)value).Length > column.MaxLength)
                    throw new MaxLengthViolationException("The value \"{0}\" exceeds the max length of {1} for column \"{2}\".",
                        value, column.MaxLength, column.Name);

                Cell cell = new Cell()
                {
                    CellReference = colRef,
                    CellValue = new CellValue(),
                    StyleIndex = stylesheet.GetStyleIndex(column.CellFormat)
                };

                column.Type.NullSafeSet(cell.CellValue, value, sharedStrings);
                cell.DataType = column.Type.DataType;

                row.Append(cell);

                colRef++;
            }

            sheetData.Append(row);
        }

        private static Sheet AddSheet(WorkbookPart workbookPart,
            uint sheetId, string name, out SheetData sheetData)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            var sheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            };

            return sheet;
        }
    }
}
