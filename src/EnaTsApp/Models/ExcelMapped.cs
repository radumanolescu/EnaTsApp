using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Com.Ena.Timesheet
{
    public class ExcelMapped
    {
        private readonly string _inputPath;
        private readonly string _outputPath;
        protected readonly ExcelPackage _excelPackage;

        public ExcelMapped(string inputPath, string outputPath)
        {
            _inputPath = inputPath;
            _outputPath = outputPath;

            // If the input path is empty, return early without initialization
            if (string.IsNullOrEmpty(inputPath))
            {
                return;
            }

            if (string.IsNullOrEmpty(outputPath))
            {
                outputPath = inputPath;
                _outputPath = outputPath;
            }

            // Validate paths
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("Input file not found", inputPath);
            }

            // Create output directory if it doesn't exist
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Initialize ExcelPackage
            _excelPackage = new ExcelPackage(new FileInfo(inputPath));
        }

        /// <summary>
        /// Saves the ExcelPackage to the output path specified in the constructor.
        /// </summary>
        public void SaveAs()
        {
            if (_excelPackage != null)
            {
                _excelPackage.SaveAs(new FileInfo(_outputPath));
            }
        }

        /// <summary>
        /// Searches for a cell with the specified value in the given row and returns its column index.
        /// The search is performed in the first worksheet of the Excel package.
        /// </summary>
        /// <param name="cellValue">The value to search for in the specified row.</param>
        /// <param name="excelRowId">The row number (1-based) to search within.</param>
        /// <returns>The 1-based column index of the found cell, or 0 if no matching cell is found.</returns>
        /// <remarks>
        /// This method performs a case-sensitive search and requires an exact match of the cell value.
        /// The search is limited to the first worksheet in the Excel package.
        /// </remarks>
        public int excelColumnOf(string cellValue, int excelRowId)
        {
            var worksheet = _excelPackage.Workbook.Worksheets[0];
            for (int colId = 1; colId <= worksheet.Dimension.End.Column; colId++)
            {
                var cell = worksheet.Cells[excelRowId, colId];
                if (cell != null && cell.Value != null && cell.Value.ToString() == cellValue)
                {
                    return colId;
                }
            }
            return 0;
        }

        /// <summary>
        /// Searches for a cell with the specified value in the given column and returns its row index.
        /// The search is performed in the first worksheet of the Excel package.
        /// </summary>
        /// <param name="cellValue">The value to search for in the specified column.</param>
        /// <param name="excelColumnId">The column number (1-based) to search within.</param>
        /// <returns>The 1-based row index of the found cell, or 0 if no matching cell is found.</returns>
        /// <remarks>
        /// This method performs a case-sensitive search and requires an exact match of the cell value.
        /// The search is limited to the first worksheet in the Excel package.
        /// </remarks>
        public int excelRowOf(string cellValue, int excelColumnId)
        {
            var worksheet = _excelPackage.Workbook.Worksheets[0];
            var column = worksheet.Column(excelColumnId);
            for (int rowId = 1; rowId <= worksheet.Dimension.End.Row; rowId++)
            {
                var cell = worksheet.Cells[rowId, excelColumnId];
                if (cell != null && cell.Value != null && cell.Value.ToString() == cellValue)
                {
                    return rowId;
                }
            }
            return 0;
        }

        public string CellValue(ExcelRangeBase cell)
        {
            if (cell == null) return string.Empty;

            switch (cell.Value)
            {
                case string stringValue:
                    return stringValue;
                case double numericValue:
                    return numericValue.ToString();
                case bool booleanValue:
                    return booleanValue.ToString();
                case null:
                    return string.Empty;
                default:
                    return cell.Text;
            }
        }

    public string OutputPath => _outputPath;
    }

}
