using System.Collections;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImports.Core
{
    internal class ExcelExporter
    {
        public void Export(WorkbookConfiguration workbookConfig, object workbookSource, Stream output)
        {
            var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = workbookConfig.StylesheetProvider.Stylesheet;

            var sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            uint sheetId = 1;

            foreach (var worksheetConfig in workbookConfig)
            {
                SheetData sheetData;
                var sheet = AddSheet(workbookPart, sheetId, worksheetConfig.SheetName, out sheetData);

                var headerRow = CreateHeaderRow(worksheetConfig, sheetData);
                sheetData.Append(headerRow);

                ICollection boundMember = worksheetConfig.GetMember(workbookSource);
                ExportData(boundMember, worksheetConfig, sheetData, workbookConfig.StylesheetProvider);

                sheets.Append(sheet);
                sheetId++;
            }

            workbookPart.Workbook.Save();
            document.Close();
        }

        private Row CreateHeaderRow(WorksheetConfiguration worksheetConfig, SheetData sheetData)
        {
            var row = new Row();

            ColumnReference colRef = "A";
            foreach (var column in worksheetConfig)
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

        private void ExportData(ICollection collection, WorksheetConfiguration worksheetConfig, SheetData sheetData,
            IStylesheetProvider stylesheet)
        {
            foreach (object item in collection)
            {
                Row row = new Row();
                ColumnReference colRef = "A";

                foreach (var column in worksheetConfig)
                {
                    // TODO [ccb] Stronger type binding would be good here.
                    // Ex: column.Type.GetValue -> Type would be an NH style IType
                    // Would need to return something that included the value and the
                    // CellValues enum for excel formatting/"typing".

                    CellBinder binder = column.GetValue(item);

                    Cell cell = new Cell()
                    {
                        CellReference = colRef,
                        CellValue = new CellValue(binder.Value),
                        DataType = new EnumValue<CellValues>(binder.CellType),
                        StyleIndex = stylesheet.GetStyleIndex(column.CellFormat)
                    };

                    row.Append(cell);

                    colRef++;
                }

                sheetData.Append(row);
            }
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
