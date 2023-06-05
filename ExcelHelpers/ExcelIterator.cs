﻿using System;
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
    public class ExcelIterator : IEnumerator<object?[]>, IEnumerable<object?[]>
    {
        #region Fields

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

        public ExcelIterator(SpreadsheetDocument document)
        {
            _document = document;
            LoadWorkbookData();
        }

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

        public string? WorksheetName { get; set; }

        public uint MinRow
        {
            get
            {
                return _fromRow;
            }

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

        public uint MaxRow
        {
            get
            {
                return _toRow;
            }

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

        public uint RowIndex { get => _rowsEnumerator?.Current.RowIndex!.Value ?? 0; }

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
            if (cell.CellValue == null) // cell not found
            {
                return null;
            }
            else if (cell.DataType is null) // text or string or date
            {
                if (Decimal.TryParse(cell.CellValue.Text, out Decimal val))
                {
                    return val;
                }
                else if (cell.CellValue.Text.Trim() == string.Empty)
                {
                    return null;
                }
                else
                {
                    return cell.CellValue.Text;
                }
            }
            else // Typed cell
            {
                return cell.DataType.Value switch
                {
                    Spreadsheet.CellValues.Boolean => int.Parse(cell.CellValue.Text),
                    Spreadsheet.CellValues.Error => null,
                    Spreadsheet.CellValues.SharedString => (string)_sharedStrings!.ElementAt(int.Parse(cell.CellValue.Text)).InnerText,
                    Spreadsheet.CellValues.String => cell.CellValue.Text.Trim() == string.Empty ? null : cell.CellValue.Text,
                    _ => null,
                };
            }
        }

        #endregion

    }
}
