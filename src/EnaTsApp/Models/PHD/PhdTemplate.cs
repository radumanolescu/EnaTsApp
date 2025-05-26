using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Com.Ena.Timesheet.Phd;
using Ena.Timesheet.Ena;

namespace Com.Ena.Timesheet.Phd
{
    public class PhdTemplate
    {
        private List<PhdTemplateEntry> entries;
        private string yearMonth; // yyyyMM
        private byte[] xlsxBytes;

        public static readonly int colOffset = 1;

        public PhdTemplate(string yearMonth, List<PhdTemplateEntry> entries)
        {
            this.yearMonth = yearMonth;
            this.entries = entries;
        }

        public PhdTemplate(string yearMonth, Stream inputStream)
        {
            this.yearMonth = yearMonth;
            using (var ms = new MemoryStream())
            {
                inputStream.CopyTo(ms);
                this.xlsxBytes = ms.ToArray();
                var parser = new Parser();
                this.entries = parser.ParseBytes(xlsxBytes);
            }
        }

        public PhdTemplate(string yearMonth, byte[] bytes)
        {
            this.yearMonth = yearMonth;
            var parser = new Parser();
            this.entries = parser.ParseBytes(bytes);
            this.xlsxBytes = bytes;
        }

        public PhdTemplate(string yearMonth, FileInfo phdTemplateFile)
        {
            this.yearMonth = yearMonth;
            using (var inputStream = phdTemplateFile.OpenRead())
            {
                var parser = new Parser();
                this.entries = parser.ParseEntries(inputStream);
                using (var ms = new MemoryStream())
                {
                    inputStream.Seek(0, SeekOrigin.Begin);
                    inputStream.CopyTo(ms);
                    this.xlsxBytes = ms.ToArray();
                }
            }
        }

        public List<PhdTemplateEntry> GetEntries()
        {
            return entries;
        }

        public byte[] Dropdowns()
        {
            var sb = new StringBuilder();
            foreach (var clientTask in ClientTasks())
            {
                sb.Append(clientTask).Append("\r\n");
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        public List<string> ClientTasks()
        {
            return entries
                .Where(entry => entry.GetTask() != "TASK")
                .Select(entry => entry.ClientHashTask())
                .ToList();
        }

        public HashSet<string> ClientTaskSet()
        {
            return new HashSet<string>(ClientTasks());
        }

        public void SetEntries(List<PhdTemplateEntry> entries)
        {
            this.entries = entries;
        }

        public string GetYearMonth()
        {
            return yearMonth;
        }

        public void SetYearMonth(string yearMonth)
        {
            this.yearMonth = yearMonth;
        }

        public Dictionary<int, double> TotalHoursByDay { get; set; } = new Dictionary<int, double>();

        public double TotalHours()
        {
            return entries.Sum(entry => entry.TotalHours());
        }

        public byte[] GetXlsxBytes()
        {
            return xlsxBytes;
        }

        public void SetXlsxBytes(byte[] xlsxBytes)
        {
            this.xlsxBytes = xlsxBytes;
        }

        public void Update(EnaTimesheet enaTimesheet)
        {
            CheckTasks(enaTimesheet);
            var enaEffort = enaTimesheet.TotalHoursByClientTaskDay();
            foreach (var entry in entries)
            {
                if (entry.Day.HasValue)
                {
                    if (enaEffort.TryGetValue(entry.ClientHashTask(), out var enaEntry))
                    {
                        var effort = new Dictionary<int?, double>();
                        foreach (var dayEffort in enaEntry)
                        {
                            effort[dayEffort.Key] = dayEffort.Value;
                        }
                        entry.SetEffort(effort);
                    }
                }
            }
            foreach (var phdEntry in entries)
            {
                if (enaEffort.TryGetValue(phdEntry.ClientHashTask(), out var enaEntry))
                {
                    var effort = new Dictionary<int?, double>();
                    foreach (var dayEffort in enaEntry)
                    {
                        effort[dayEffort.Key] = dayEffort.Value;
                    }
                    phdEntry.SetEffort(effort);
                }
            }
            double enaTotalHours = enaTimesheet.TotalHours();
            double phdTotalHours = TotalHours();
            if (enaTotalHours != phdTotalHours)
            {
                PrintEntries();
                throw new Exception($"Total hours mismatch: ENA={enaTotalHours} PHD={phdTotalHours}");
            }
            UpdateXlsx(0);
        }

        private void PrintEntries()
        {
            foreach (var entry in entries)
            {
                Console.WriteLine(entry.ClientTaskEffort());
            }
        }

        public void CheckTasks(EnaTimesheet enaTimesheet)
        {
            var clientTaskSet = ClientTaskSet();
            foreach (var enaEntry in enaTimesheet.GetEntries())
            {
                string matchKey = enaEntry.ProjectActivity();
                if (!clientTaskSet.Contains(matchKey))
                {
                    throw new Exception("Unknown ENA task: " + matchKey);
                }
            }
        }

        public void UpdateXlsx(int index)
        {
            if (xlsxBytes == null) return;
            using (var inputStream = new MemoryStream(xlsxBytes))
            {
                IWorkbook workbook = WorkbookFactory.Create(inputStream);
                ISheet sheet = workbook.GetSheetAt(index);
                int rowId = 0;
                int numEntries = entries.Count;
                foreach (IRow row in sheet)
                {
                    if (rowId >= numEntries) break;
                    EraseEffort(rowId, row);
                    var entry = entries[rowId];
                    foreach (var dayEffort in entry.GetEffort())
                    {
                        int day = dayEffort.Key.Value;
                        double effort = dayEffort.Value;
                        ICell cell = row.GetCell(colOffset + day) ?? row.CreateCell(colOffset + day);
                        cell.SetCellValue(effort);
                    }
                    rowId++;
                }
                workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
                using (var outputStream = new MemoryStream())
                {
                    workbook.Write(outputStream);
                    xlsxBytes = outputStream.ToArray();
                }
            }
        }

        // Erase the effort for the row
        private void EraseEffort(int rowId, IRow row)
        {
            // Skip the first row (header)
            if (rowId > 0)
            {
                // Erase the row to make sure we don't have any old data
                for (int colId = colOffset + 1; colId <= colOffset + 31; colId++)
                {
                    ICell cell = row.GetCell(colId);
                    if (cell != null)
                    {
                        cell.SetCellType(CellType.Blank);
                    }
                }
            }
        }

        public void WriteFile(FileInfo output)
        {
            File.WriteAllBytes(output.FullName, xlsxBytes);
        }
    }
}