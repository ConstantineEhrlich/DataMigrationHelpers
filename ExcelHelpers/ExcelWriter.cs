using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;



namespace ExcelHandlers
{
    public class ExcelWriter: IDisposable
    {
        #region Fields

        private readonly string _exportPath;
        private SpreadsheetDocument _document;
        private WorkbookPart _workbook;
        private Dictionary<string, uint> _styles;

        #endregion



        #region Constructors

        public ExcelWriter(string exportPath)
        {
            _exportPath = exportPath;
            _document = SpreadsheetDocument.Create(_exportPath, SpreadsheetDocumentType.Workbook);
            _workbook = _document.AddWorkbookPart();
            _workbook.Workbook = new Workbook();
            _workbook.Workbook.AppendChild<Sheets>(new Sheets());
            _workbook.AddNewPart<WorkbookStylesPart>();
            _workbook.WorkbookStylesPart!.Stylesheet = CreateExcelStylesheet.CreateStylesheet(out _styles);
            _workbook.Workbook.Save();
        }


        #endregion



        #region Properties

        public SpreadsheetDocument Document { get => _document; }

        #endregion



        #region Indexers
        #endregion



        #region Public Methods

        public void WriteData(string sheetName, IDataReader data)
        {
            WorksheetPart ws = AddWorksheet(_workbook, sheetName);

            // Read header row
            data.Read();
            string[] headers = new string[data.FieldCount];
            for (int i = 0; i < data.FieldCount; i++)
            {
                headers[i] = data.GetName(i);
            }

            // Write header row
            int rowIndex = 1;
            ws.Worksheet.GetFirstChild<SheetData>()!.AppendChild(MakeRow(headers, rowIndex));


            do
            {
                rowIndex++;
                object[] vals = new object[data.FieldCount];
                data.GetValues(vals);
                ws.Worksheet.GetFirstChild<SheetData>()!.AppendChild(MakeRow(vals, rowIndex));
            }
            while (data.Read());

            ws.Worksheet.Save();
            _workbook.Workbook.Save();
        }

        public void Dispose()
        {
            _document.Save();
            _document.Close();
            _document.Dispose();
            GC.SuppressFinalize(this);
        }

        #endregion



        #region Private Methods

        private WorksheetPart AddWorksheet(WorkbookPart workbook, string sheetName)
        { 
            // Add new worksheet to the workbook
            WorksheetPart ws = workbook.AddNewPart<WorksheetPart>();
            ws.Worksheet = new Spreadsheet.Worksheet(new SheetData());

            // Get maximum sheet id
            int sheetId = workbook.Workbook.Sheets?.Count() + 1 ?? 1;

            // Add new sheet to the workbook as part of the document
            Sheet sheet = new()
            {
                Id = this._workbook.GetIdOfPart(ws),
                SheetId = (uint)sheetId,
                Name = sheetName,
            };
            workbook.Workbook.Sheets!.Append(sheet);

            return ws;
        }


        private Row MakeRow(object?[] values, int rowIndex)
        {
            Row row = new();
            for (int i = 0; i < values.Length; i++)
            {
                Cell cell = MakeCell(values[i]);
                cell.CellReference = ExcelExtensions.IndexToLetter(i + 1) + rowIndex.ToString();
                // Make header row bold
                if (rowIndex == 1)
                {
                    cell.StyleIndex = _styles["Header"];
                }

                row.AppendChild(cell);
            }
            return row;
        }


        private Cell MakeCell(object? val)
        {
            Cell cell = new();

            if (val is null || val == DBNull.Value)
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(string.Empty);
                return cell;
            }

            // Get value type, or underlying type since val can be nullable
            Type valType = Nullable.GetUnderlyingType(val.GetType()) ?? val.GetType();

            switch (valType.Name)
            {
                case "DateTime":
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(Convert.ToDateTime(val).ToOADate());
                    cell.StyleIndex = _styles["Date"];
                    return cell;
                case "Int16":
                case "Int32":
                case "Int64":
                case "Single":
                case "Double":
                case "Decimal":
                    cell.DataType = CellValues.Number;
                    cell.StyleIndex = _styles["Number"];
                    cell.CellValue = new CellValue(Convert.ToDecimal(val));
                    return cell;
                case "String":
                default:
                    cell.DataType = CellValues.String;
                    cell.StyleIndex = _styles["Default"];
                    cell.CellValue = new CellValue(val.ToString() ?? string.Empty);
                    return cell;
            }


        }


        #endregion
    }
}
