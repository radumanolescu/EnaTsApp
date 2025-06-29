using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace Com.Ena.Timesheet.Xl
{
    public class ExcelParser
    {
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
            var data = new List<List<string>>();

            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    var workbook = WorkbookFactory.Create(fs);
                    var sheet = workbook.GetSheetAt(SHEET_INDEX);
                    if (sheet == null)
                    {
                        throw new InvalidOperationException("No worksheets found in the Excel file.");
                    }

                    var rowCount = sheet.LastRowNum + 1;
                    if (rowCount == 0)
                    {
                        throw new InvalidOperationException("The Excel file appears to be empty.");
                    }

                    for (int row = 0; row < rowCount; row++)
                    {
                        var rowData = new List<string>();
                        var excelRow = sheet.GetRow(row);
                        if (excelRow != null)
                        {
                            var colCount = excelRow.LastCellNum;
                            for (int col = 0; col < colCount; col++)
                            {
                                var cell = excelRow.GetCell(col);
                                if (cell == null)
                                {
                                    rowData.Add(string.Empty);
                                }
                                else
                                {
                                    switch (cell.CellType)
                                    {
                                        case CellType.String:
                                            rowData.Add(cell.StringCellValue);
                                            break;
                                        case CellType.Numeric:
                                            rowData.Add(cell.NumericCellValue.ToString());
                                            break;
                                        case CellType.Boolean:
                                            rowData.Add(cell.BooleanCellValue.ToString());
                                            break;
                                        case CellType.Formula:
                                            var evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                                            var cellValue = evaluator.Evaluate(cell);
                                            string formulaResult;
                                            switch (cellValue.CellType)
                                            {
                                                case CellType.String:
                                                    formulaResult = cellValue.StringValue;
                                                    break;
                                                case CellType.Numeric:
                                                    formulaResult = cellValue.NumberValue.ToString();
                                                    break;
                                                case CellType.Boolean:
                                                    formulaResult = cellValue.BooleanValue.ToString();
                                                    break;
                                                default:
                                                    formulaResult = string.Empty;
                                                    break;
                                            }
                                            rowData.Add(formulaResult);
                                            break;
                                        default:
                                            rowData.Add(string.Empty);
                                            break;
                                    }
                                }
                            }
                        }
                        data.Add(rowData);
                    }

                    return data;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error parsing Excel file: {ex.Message}", ex);
            }
        }

        public Dictionary<int, List<string>> ReadWorkbook(Stream inputStream, int index)
        {
            IWorkbook workbook = WorkbookFactory.Create(inputStream);
            return ReadWorkbook(workbook, index);
        }

        public Dictionary<int, List<string>> ReadWorkbook(IWorkbook workbook, int index)
        {
            return ReadWorksheet(workbook.GetSheetAt(index));
        }

        public Dictionary<int, List<string>> ReadWorksheet(ISheet sheet)
        {
            var data = new Dictionary<int, List<string>>();
            int i = 0;
            foreach (IRow row in sheet)
            {
                data[i] = StringValues(row);
                i++;
            }
            return data;
        }

        public List<string> StringValues(IRow row)
        {
            var cells = row.Cells;
            var rowValues = new string[cells.Count];
            for (int i = 0; i < cells.Count; i++)
            {
                rowValues[i] = cells[i].ToString();
            }
            return rowValues.ToList();
        }
    }
}