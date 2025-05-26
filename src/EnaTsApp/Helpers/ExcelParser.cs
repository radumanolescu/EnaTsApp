using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace Com.Ena.Timesheet.Xl
{
    public class ExcelParser
    {
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