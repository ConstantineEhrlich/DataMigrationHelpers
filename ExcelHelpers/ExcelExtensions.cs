using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelpers
{
    public static class ExcelExtensions
    {
        /// <summary>
        /// Converts Excel column literal to index integer.
        /// </summary>
        /// <param name="columnLiteral">Excel-style column literal, from "A" to "XFD", letters only</param>
        /// <returns>1-based index of the column</returns>
        /// <exception cref="System.ArgumentNullException">Thrown when columnLiteral is empty or null</exception>
        /// <exception cref="System.ArgumentException">Thrown when columnLiteral contains of illegal characters</exception>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when columnLiteral is out of Excel column boundaries</exception>
        public static int LetterToIndex(string columnLiteral)
        {
            if (string.IsNullOrWhiteSpace(columnLiteral))
                throw new ArgumentNullException(columnLiteral);

            if (columnLiteral.Length > 3 | (columnLiteral.Length == 3 & columnLiteral.CompareTo("XFD") > 0))
                throw new ArgumentOutOfRangeException(paramName: columnLiteral, message: $"Column {columnLiteral} is not valid Excel column");

            if (columnLiteral.Any(c => !char.IsLetter(c)))
                throw new ArgumentException(paramName: columnLiteral, message: $"Column address must contain only letters");

            columnLiteral = columnLiteral.ToUpperInvariant();
            int idx = 0;
            for (int i = 0; i < columnLiteral.Length; i++)
            {
                idx *= 26;
                idx += (columnLiteral[i] - 'A' + 1);
            }
            return idx;
        }

        /// <summary>
        /// Calculates column literal for 1-based index;
        /// </summary>
        /// <param name="index">1-based index of the column</param>
        /// <returns>Column literal string</returns>
        /// <exception cref="System.ArgumentOutOfRangeException">Thrown when provided index is out of range</exception>
        public static string IndexToLetter(int index)
        {
            string columnName = string.Empty;

            if (index <= 0 | index > 16384)
                throw new ArgumentOutOfRangeException(nameof(index), $"Column index {index} out of range!");

            while (index > 0)
            {
                int remainder = (index - 1) % 26;
                columnName = Convert.ToChar('A' + remainder) + columnName;
                index = (index - remainder) / 26;
            }
            
            
            return columnName;
        }


        public static string CellColumnLiteral(string cellAddress)
        {
            char split = cellAddress.FirstOrDefault(chr => char.IsDigit(chr));
            if (split == default(char))
                throw new ArgumentException($"Address {cellAddress} is not correct Excel cell address");

            return cellAddress.Substring(0, cellAddress.IndexOf(split));
        }


        public static int ColumnIndex(this DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            if (cell.CellReference is null || cell.CellReference.Value is null)
                throw new NullReferenceException($"Cell is null");

            return LetterToIndex(CellColumnLiteral(cell.CellReference.Value));
        }

        public static string ColumnLiteral(this DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            if (cell.CellReference is null || cell.CellReference.Value is null)
                throw new NullReferenceException($"Cell is null");

            return CellColumnLiteral(cell.CellReference.Value);
        }

    }
}
