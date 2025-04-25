using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelHelpers
{
    /// <summary>
    /// Represents an iterator over Excel rows.
    /// </summary>
    public class ExcelIterator : IEnumerator<object?[]>, IEnumerable<object?[]>
    {
        #region Fields

        private bool _use1904DateSystem = false;
        private IEnumerable<Spreadsheet.Row>? _rows;
        private IEnumerator<Spreadsheet.Row>? _rowsEnumerator;

        private SpreadsheetDocument _document;
        private WorkbookPart? _workbook;
        private Spreadsheet.SharedStringTable? _sharedStrings;

        private uint _fromRow = 0;
        private uint _toRow = 0;
        private uint _fromCol = 0;
        private uint _toCol = 0;

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the ExcelIterator class using the specified SpreadsheetDocument.
        /// </summary>
        /// <param name="document">The <see cref="SpreadsheetDocument"/> object representing the Excel document to read from.</param>
        public ExcelIterator(SpreadsheetDocument document)
        {
            _document = document;
            LoadWorkbookData();
        }

        /// <summary>
        /// Initializes a new instance of the ExcelIterator class using the specified file path.
        /// </summary>
        /// <param name="excelFilePath">The path of the Excel file to read from.</param>
        public ExcelIterator(string excelFilePath) : this(SpreadsheetDocument.Open(excelFilePath, false))
        {

        }

        #endregion



        #region Properties

        private IEnumerator<Spreadsheet.Row> RowsEnumerator
        {
            get
            {
                if (_rows is null)
                {
                    LoadWorksheetRows();
                    SelectColumns();
                    SelectRows();
                    _rowsEnumerator = _rows!.GetEnumerator();
                }
                return _rowsEnumerator!;
            }
        }

        object IEnumerator.Current { get => Current; }

        public object?[] Current { get => RowToArray(RowsEnumerator.Current.Elements<Spreadsheet.Cell>().ToArray()); }

        /// <summary>
        /// Gets or sets the name of the worksheet to be read.
        /// </summary>
        public string? WorksheetName { get; set; }

        /// <summary>
        /// Gets or sets the minimum row index for iteration. Must be a valid Excel row index and not be greater than MaxRow.
        /// </summary>
        /// <exception cref="IndexOutOfRangeException">Thrown when the value is not a correct Excel row index or when it is greater than MaxRow.</exception>
        public uint MinRow
        {
            get => _fromRow;

            set
            {
                if (value == 0 | value > 1048576)
                {
                    throw new IndexOutOfRangeException($"{value} is not correct Excel Row");
                }
                else if (_toRow != 0 && value < _toRow)
                {
                    throw new IndexOutOfRangeException($"MinRow value ({value}) cannot be greater than MaxRow ({_toRow})");
                }
                else
                {
                    _fromRow = value;
                    
                }
            }
        }

        /// <summary>
        /// Gets or sets the maximum row index for iteration. Must be a valid Excel row index and not be less than MinRow.
        /// </summary>
        /// <exception cref="IndexOutOfRangeException">Thrown when the value is not a correct Excel row index or when it is less than MinRow.</exception>
        public uint MaxRow
        {
            get => _toRow;

            set
            {
                if (value == 0 | value > 1048576)
                {
                    throw new IndexOutOfRangeException($"{value} is not correct Excel Row");
                }
                else if (_fromRow != 0 && value < _fromRow)
                {
                    throw new IndexOutOfRangeException($"ToRow value ({value}) cannot be smaller than FromRow ({_fromRow})");
                }
                else
                {
                    _toRow = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the minimum column index for iteration. Must be a valid Excel column index and not be greater than MaxCol.
        /// </summary>
        /// <exception cref="IndexOutOfRangeException">Thrown when the value is not a correct Excel column index or when it is greater than MaxCol.</exception>
        public uint MinCol
        {
            get
            {
                return _fromCol;
            }
            set
            {
                if (value == 0 | value > 16384)
                {
                    throw new IndexOutOfRangeException($"{value} is not correct Excel Column");
                }
                else if (_toCol != 0 && value > _toCol)
                {
                    throw new IndexOutOfRangeException($"MinCol value ({value}) cannot be greater than MaxCol ({_toRow})");
                }
                else
                {
                    _fromCol = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the maximum column index for iteration. Must be a valid Excel column index and not be less than MinCol.
        /// </summary>
        /// <exception cref="IndexOutOfRangeException">Thrown when the value is not a correct Excel column index or when it is less than MinCol.</exception>
        public uint MaxCol
        {
            get
            {
                return _toCol;
            }
            set
            {
                if (value == 0 | value > 16384)
                {
                    throw new IndexOutOfRangeException($"{value} is not correct Excel Column");
                }
                else if (_fromCol != 0 && value < _fromCol)
                {
                    throw new IndexOutOfRangeException($"MaxCol value ({value}) cannot be smaller than MinCol ({_fromRow})");
                }
                else
                {
                    _toCol = value;
                }
            }
        }

        /// <summary>
        /// Gets the index of the current row.
        /// </summary>
        public uint RowIndex => _rowsEnumerator?.Current.RowIndex!.Value ?? 0;

        #endregion



        #region Indexers
        #endregion


        #region Public Methods

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public IEnumerator<object?[]> GetEnumerator()
        {
            if (_rows is not null)
            {
                _rows = null;
            }
            return this;
        }

        public bool MoveNext()
        {
            if (!RowsEnumerator.MoveNext())
            {
                _rows = null;
                return false;
            }
            else
            {
                return true;
            }
        }
            
        public void Reset()
        {
            _rows = null;
        }
        
        public void Dispose()
        {
            if (_document is not null)
            {
                _document.Dispose();
            }
            GC.SuppressFinalize(this);
        }

        #endregion



        #region Private Methods

        /// <summary>
        /// Loads _workbook and _sharedStrings
        /// </summary>
        private void LoadWorkbookData()
        {
            _workbook = _document.WorkbookPart
                ?? throw new NullReferenceException($"File does not contain proper Workbook Part");

            _sharedStrings = _workbook.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable
                ?? throw new NullReferenceException($"Shared strings part is missing in the file");
            
            // Check if 1904 date system is in use
            var wbProps = _workbook.Workbook.Descendants<Spreadsheet.WorkbookProperties>().FirstOrDefault();
            if (wbProps is not null && wbProps.Date1904 is not null)
                _use1904DateSystem = wbProps.Date1904.Value;
            {
                
            }

        }

        /// <summary>
        /// Loads _rows
        /// </summary>
        private void LoadWorksheetRows()
        {
            WorksheetPart worksheet;

            // Check if there is any sheet with name equals to the property WorksheetName
            Spreadsheet.Sheet? sheet = _workbook!.Workbook
                                         .Descendants<Spreadsheet.Sheet>()
                                         .FirstOrDefault(sheet => sheet.Name == WorksheetName);


            // If the sheet is not found, take the first sheet, else take the sheet with sheet's Id
            if (sheet is null)
            {
                worksheet = _workbook.WorksheetParts.First();
            }
            else
            {
                worksheet = (WorksheetPart)_workbook.GetPartById(sheet.Id!);
            }


            // When the worksheet is found, get sheet data and sheet rows from it
            Spreadsheet.SheetData sheetData = worksheet.Worksheet.Elements<Spreadsheet.SheetData>().First();
            _rows = sheetData.Elements<Spreadsheet.Row>();
        }

        /// <summary>
        /// Handles MinCol and MaxCol properties
        /// </summary>
        private void SelectColumns()
        {
            // If the maximum column is not set, set it to max span
            if (_toCol == 0)
            {
                uint maxCol = _rows!
                    .Select(row => row.Spans!.InnerText!.Split(':').Last())
                    .Select(rowSpan => uint.Parse(rowSpan))
                    .Max();

                if (maxCol >= _fromCol)
                {
                    _toCol = maxCol;
                }
            }

            // If the minimum column is not set, set it to 1
            if (_fromCol == 0)
                _fromCol = 1;
        }

        /// <summary>
        /// Handles MinRow and MaxRow properties
        /// </summary>
        private void SelectRows()
        {
            // Trim the rows according to the properties
            if (_fromRow != 0)
                _rows = _rows!.Where(row => row.RowIndex! >= _fromRow);

            if (_toRow != 0)
                _rows = _rows!.Where(row => row.RowIndex! <= _toRow);
        }

        /// <summary>
        /// Converts an array of Spreadsheet.Cell into the array on values
        /// </summary>
        /// <param name="row">Array of Spreadsheet.Cell objects</param>
        /// <returns>Array of values</returns>
        private object?[] RowToArray(Spreadsheet.Cell[] row)
        {
            object?[] result = new object?[_toCol - _fromCol + 1];

            for (int i = 0; i < result.Length; i++)
            {
                Spreadsheet.Cell? cell = row.FirstOrDefault(cell => cell.ColumnIndex() == i + MinCol);
                if (cell is null)
                {
                    result[i] = null;
                }
                else
                {
                    result[i] = GetCellValue(cell);
                }
            }
            return result;
        }

        /// <summary>
        /// Analyzes the cell and returns its value
        /// </summary>
        /// <param name="cell">Excel cell</param>
        /// <returns>Cell value or null if the cell is empty</returns>
        private object? GetCellValue(Spreadsheet.Cell cell)
        {
            // If no value, return null
            if (cell.CellValue == null)
                return null;

            string cellText = cell.CellValue.Text;
            if (cell.DataType is null)
            {
                // Check the style
                var styleIndex = cell.StyleIndex?.Value;
                if (styleIndex is not null && _workbook?.WorkbookStylesPart?.Stylesheet?.CellFormats is not null)
                {
                    // Parse as date if style checks out
                    var cellFormat = (Spreadsheet.CellFormat)_workbook.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int)styleIndex);
                    var formatId = cellFormat.NumberFormatId?.Value;
                    if (IsDateFormat(formatId) && double.TryParse(cellText, out var numericDate))
                    {
                        double offset = _use1904DateSystem ? 1462 : 0;
                        return DateTime.FromOADate(numericDate + offset);
                    }
                }
                
                if (decimal.TryParse(cell.CellValue.Text, out decimal decimalVal))
                    return decimalVal;
                
                if (string.IsNullOrWhiteSpace(cell.CellValue.Text))
                    return null;
                
                return cell.CellValue.Text;
            }
            else // Typed cell
            {
                
                if (cell.DataType.Value == Spreadsheet.CellValues.Boolean)
                    return int.Parse(cell.CellValue.Text);
                else if (cell.DataType.Value == Spreadsheet.CellValues.Error)
                    return null;
                else if (cell.DataType.Value == Spreadsheet.CellValues.SharedString)
                    return _sharedStrings!.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
                else if (cell.DataType.Value == Spreadsheet.CellValues.String)
                    return cell.CellValue.Text.Trim() == string.Empty ? null : cell.CellValue.Text;
                else
                    return null;

            }
        }

        private bool IsDateFormat(uint? numberFormatId)
        {
            if (numberFormatId == null)
                return false;

            var dateIds = new HashSet<uint>() { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };
            return dateIds.Contains(numberFormatId.Value);
        }

        #endregion

    }
}
