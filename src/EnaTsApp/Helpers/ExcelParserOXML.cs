using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Com.Ena.Timesheet.Xl
{
    /// <summary>
    /// Parses Excel files using the EPPlus library.
    /// This implementation is functionally equivalent to ExcelParser but uses EPPlus instead of NPOI.
    /// </summary>
    public class ExcelParser
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        public ExcelParser()
        {
            // Constructor is guaranteed to be called on every instantiation
            // Set the license context
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        }
        private const int SHEET_INDEX = 0;

        /// <summary>
        /// Parses an Excel file and returns its contents as a list of lists, where each inner list represents a row.
        /// This method is agnostic to the Excel file format and can be used to parse any Excel file.
        /// It should not be customized with assumptions about any specific Excel file format.
        /// </summary>
        /// <param name="filePath">The path to the Excel file to parse.</param>
        /// <returns>A list of lists containing the Excel data, where each inner list represents a row and each string represents a cell value.</returns>
        /// <exception cref="System.ArgumentNullException">Thrown when the filePath is null.</exception>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the specified file does not exist.</exception>
        /// <exception cref="System.InvalidOperationException">Thrown when no worksheets are found in the Excel file or the file appears to be empty.</exception>
        public List<List<string>> ParseExcelFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified Excel file was not found.", filePath);

            var data = new List<List<string>>();

            var fileInfo = new FileInfo(filePath);
            using (var package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("No worksheets found in the Excel file.");
                }

                var worksheet = package.Workbook.Worksheets[SHEET_INDEX];
                var dimension = worksheet.Dimension;
                
                if (dimension == null)
                {
                    throw new InvalidOperationException("The Excel file appears to be empty or corrupted.");
                }

                // Get the actual used range
                var start = dimension.Start;
                var end = dimension.End;

                for (int row = 1; row <= end.Row; row++)
                {
                    var rowData = new List<string>();
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell == null || cell.Value == null)
                        {
                            rowData.Add(string.Empty);
                        }
                        else
                        {
                            // Handle different value types
                            var value = cell.Value;
                            string stringValue;

                            if (value is DateTime dateTimeValue)
                            {
                                stringValue = dateTimeValue.ToString("G");
                            }
                            else if (value is bool boolValue)
                            {
                                stringValue = boolValue.ToString().ToLower();
                            }
                            else if (value is double || value is float || value is decimal || value is int || value is long)
                            {
                                // Preserve numeric precision
                                stringValue = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                stringValue = value.ToString();
                            }

                            rowData.Add(stringValue);
                        }
                    }
                    data.Add(rowData);
                }
            }

            return data;
        }
    }
}
