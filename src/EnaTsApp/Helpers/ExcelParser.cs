using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace Com.Ena.Timesheet.Xl
{
    public class ExcelParser
    {
        private const int SHEET_INDEX = 0;

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