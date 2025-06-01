using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Ena;
using Com.Ena.Timesheet;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;

namespace Com.Ena.Timesheet.Phd
{
    public class PhdTemplate : ExcelMapped
    {
        private readonly ILogger<PhdTemplate> _logger;
        private List<PhdTemplateEntry> entries;
        private string yearMonth; // yyyyMM
        private IWorkbook workbook;

        private static readonly IServiceProvider _serviceProvider = new ServiceCollection()
            .AddLogging(builder => builder.AddConsole())
            .BuildServiceProvider();

        private static ILogger<T> GetLogger<T>()
        {
            return _serviceProvider.GetRequiredService<ILogger<T>>();
        }

        public string YearMonth => yearMonth;
        public static readonly int colOffset = 1;

        public PhdTemplate(string yearMonth, List<PhdTemplateEntry> entries, string inputPath, string outputPath) 
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<PhdTemplate>();
            _logger.LogInformation($"Creating PhdTemplate for {yearMonth}");
            this.yearMonth = yearMonth;
            this.entries = entries;
            this.workbook = new XSSFWorkbook();
        }

        public PhdTemplate(string yearMonth, List<List<string>>? templateData, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<PhdTemplate>();
            _logger.LogInformation($"Creating PhdTemplate from data for {yearMonth}");
            this.yearMonth = yearMonth;

            this.entries = new List<PhdTemplateEntry>();
            for (int i = 0; i < templateData.Count; i++)
            {
                var entry = new PhdTemplateEntry(i, templateData[i][0], templateData[i][1]);
                this.entries.Add(entry);
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
                        var effort = new Dictionary<int, double>();
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
                    var effort = new Dictionary<int, double>();
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

        // ToDo: fix this
        public void UpdateXlsx(int index)
        {
            using (var inputStream = new MemoryStream())
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
                        int day = dayEffort.Key;
                        double effort = dayEffort.Value;
                        ICell cell = row.GetCell(colOffset + day) ?? row.CreateCell(colOffset + day);
                        cell.SetCellValue(effort);
                    }
                }
                workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
                workbook.Write(inputStream);
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
            if (workbook == null)
            {
                workbook = new XSSFWorkbook();
            }
            
            using (var outputStream = new FileStream(output.FullName, FileMode.Create))
            {
                workbook.Write(outputStream);
            }
        }
    }
}