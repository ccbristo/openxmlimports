using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImports.Core
{
    public class DefaultStylesheetProvider : IStylesheetProvider
    {
        public NumberingFormat DateFormat { get; private set; }
        public NumberingFormat DateTimeFormat { get; private set; }
        public NumberingFormat TimeFormat { get; private set; }

        public Stylesheet Stylesheet { get; private set; }

        public uint GetStyleIndex(NumberingFormat format)
        {
            if (format == null)
                return 0;

            var formats = this.Stylesheet.Descendants<NumberingFormat>().ToList();
            int index = formats.IndexOf(format);

            if (index == -1)
                // TODO [ccb] Should this be a hard exception like below
                // or should an error policy be used?
                // If it stays a hard exception, a more specific one should be thrown.
                throw new InvalidOperationException("The requested numbering format does not belong to the stylesheet.");

            return (uint)index + 1; // excel indices are 1 based
        }

        public DefaultStylesheetProvider()
        {
            this.Stylesheet = CreateStylesheet();
        }

        private Stylesheet CreateStylesheet()
        {
            Stylesheet stylesheet = new Stylesheet();
            Fonts fonts = new Fonts();
            Font font = new Font();
            FontName fontName = new FontName();
            fontName.Val = "Calibri";
            FontSize fontSize = new FontSize();
            fontSize.Val = 11;
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);
            fonts.Count = (uint)fonts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            PatternFill patternFill;
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats cellStyleFormats = new CellStyleFormats();
            CellFormat cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 0;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellStyleFormats.Append(cellFormat);
            cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;

            uint iExcelIndex = 164;

            NumberingFormats numberFormats = new NumberingFormats();
            CellFormats cellFormats = new CellFormats();

            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 0;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormats.Append(cellFormat);

            DateTimeFormat = new NumberingFormat();
            DateTimeFormat.NumberFormatId = iExcelIndex++;
            DateTimeFormat.FormatCode = string.Format("{0} {1}",
                                            CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern,
                                            CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern)
                                            .Replace("tt", "AM/PM");
            numberFormats.Append(DateTimeFormat);
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = DateTimeFormat.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = true;
            cellFormats.Append(cellFormat);

            DateFormat = new NumberingFormat();
            DateFormat.NumberFormatId = iExcelIndex++;
            DateFormat.FormatCode = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            numberFormats.Append(DateFormat);
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = DateFormat.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = true;
            cellFormats.Append(cellFormat);

            TimeFormat = new NumberingFormat();
            TimeFormat.NumberFormatId = iExcelIndex++;
            TimeFormat.FormatCode = CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern.Replace("tt", "AM/PM");
            numberFormats.Append(TimeFormat);
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = TimeFormat.NumberFormatId;
            cellFormat.FontId = 0;
            cellFormat.FillId = 0;
            cellFormat.BorderId = 0;
            cellFormat.FormatId = 0;
            cellFormat.ApplyNumberFormat = true;
            cellFormats.Append(cellFormat);

            numberFormats.Count = (uint)numberFormats.ChildElements.Count;
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            stylesheet.Append(numberFormats);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);

            CellStyles css = new CellStyles();
            CellStyle cs = new CellStyle();
            cs.Name = "Normal";
            cs.FormatId = 0;
            cs.BuiltinId = 0;
            css.Append(cs);
            css.Count = (uint)css.ChildElements.Count;
            stylesheet.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;
            stylesheet.Append(dfs);

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = "TableStyleMedium9";
            tss.DefaultPivotStyle = "PivotStyleLight16";
            stylesheet.Append(tss);

            return stylesheet;
        }
    }
}