using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandlers
{
    static class CreateExcelStylesheet
    {
        public static Stylesheet CreateStylesheet(out Dictionary<string, uint> styleMap)
        {
            Stylesheet stylesheet = new();
            styleMap = new();

            #region Numbering Formats
            NumberingFormats formats = new();
            NumberingFormat dateFormat = new()
            {
                NumberFormatId = 164,
                FormatCode = "dd-mmm-yyyy",
            };
            NumberingFormat numFormat = new()
            {
                NumberFormatId = 165,
                FormatCode = "_-* #,##0_-; * (#,##0);_-* \" - \"??_-;_-@_-",
            };
            formats.Append(dateFormat);
            formats.Append(numFormat);
            formats.Count = 2;
            stylesheet.Append(formats);
            #endregion


            #region Fonts
            Fonts fonts = new()
            { 
                KnownFonts = BooleanValue.FromBoolean(true)
            };
            Font fntDef = new()
            {
                FontSize = new FontSize() { Val = 12 },
                FontName = new FontName() { Val = "Arial Narrow" },
            };
            fonts.Append(fntDef);
            Font fnBold = new()
            {
                FontSize = new FontSize() { Val = 12 },
                FontName = new FontName() { Val = "Arial Narrow" },
                Bold = new Bold() { Val = BooleanValue.FromBoolean(true) },
            };
            fonts.Append(fnBold);
            Font fnNumbers = new()
            {
                FontSize = new FontSize() { Val = 12 },
                FontName = new FontName() { Val = "Consolas" },
            };
            fonts.Append(fnNumbers);
            stylesheet.Append(fonts);
            #endregion


            #region Fills
            Fills fills = new();
            Fill fillDef = new()
            {
                PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.None) },
            };
            Fill fillGray125 = new()
            {
                PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.Gray125) }
            };
            fills.Append(fillDef);
            fills.Append(fillGray125);
            //fills.Count = 2;
            stylesheet.Append(fills);
            #endregion


            #region Borders
            Borders borders = new();
            Border border = new()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder(),
            };
            borders.Append(border);
            //borders.Count = 1;
            stylesheet.Append(borders);
            #endregion


            #region CellStyleFormats
            CellStyleFormats cellStyleFormats = new();
            CellFormat defCellStyleFormat = new();
            cellStyleFormats.Append(defCellStyleFormat);
            //cellStyleFormats.Count = 1;
            stylesheet.Append(cellStyleFormats);
            #endregion


            #region CellFormats
            uint styleCount = 0;
            CellFormats cellFormats = new();
            CellFormat defCellFormat = new()
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
            };
            cellFormats.Append(defCellFormat);
            styleMap.Add("Default", styleCount++);

            CellFormat dateCellFormat = new()
            {
                NumberFormatId = 164,
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
            };
            cellFormats.Append(dateCellFormat);
            styleMap.Add("Date", styleCount++);

            CellFormat numCellFormat = new()
            {
                NumberFormatId = 165,
                FormatId = 0,
                FontId = 2, // Monospace font
                BorderId = 0,
                FillId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
            };
            cellFormats.Append(numCellFormat);
            styleMap.Add("Number", styleCount++);

            CellFormat headerCellFormat = new()
            {
                FormatId = 0,
                FontId = 1, // Bold font
                BorderId = 0,
                FillId = 0,
                ApplyFont = BooleanValue.FromBoolean(true),
            };
            cellFormats.Append(headerCellFormat);
            styleMap.Add("Header", styleCount++);

            //cellFormats.Count = 5;
            stylesheet.Append(cellFormats);
            #endregion

            return stylesheet;
        }
    }
}
