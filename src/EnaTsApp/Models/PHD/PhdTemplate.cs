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
using OfficeOpenXml;

namespace Com.Ena.Timesheet.Phd
{
    public class PhdTemplate : ExcelMapped
    {
        private readonly ILogger<PhdTemplate> _logger;
        private List<PhdTemplateEntry> entries;
        private string yearMonth; // yyyyMM
        private IWorkbook workbook;
        private ExcelWorksheet worksheet;

        private static readonly IServiceProvider _serviceProvider = new ServiceCollection()
            .AddLogging(builder => builder.AddConsole())
            .BuildServiceProvider();

        private static ILogger<T> GetLogger<T>()
        {
            return _serviceProvider.GetRequiredService<ILogger<T>>();
        }

        public string YearMonth => yearMonth;
        /*
            In OfficeOpenXml, the row and column indexes are 1-based:

            Row indexes: Start from 1 (not 0)
            Column indexes: Start from 1 (not 0)
            For example:

            First row is row 1
            First column is column 1
            Cell A1 is accessed as (1, 1)
            Cell B2 is accessed as (2, 2)
        */
        public static readonly int colOffset = 1;

        public PhdTemplate(string yearMonth, List<PhdTemplateEntry> entries, string inputPath, string outputPath) 
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<PhdTemplate>();
            _logger.LogInformation($"Creating PhdTemplate for {yearMonth}");
            this.yearMonth = yearMonth;
            this.entries = entries;
            workbook = WorkbookFactory.Create(inputPath);
            worksheet = _excelPackage.Workbook.Worksheets[0];
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
                string client = templateData[i][0];
                if (client == "SUM")
                    break;
                string task = templateData[i][1];
                var entry = new PhdTemplateEntry(i, client, task);
                this.entries.Add(entry);
            }
            Parser.SetProjectCodes(entries);
            Parser.CheckDupClientTask(entries);
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

        /// <summary>
        /// Updates this PHD template's entries with data from the ENA timesheet.
        /// 
        /// This method performs the following steps:
        /// 1. Validates that all ENA tasks exist in the PHD template
        /// 2. Maps hours from ENA timesheet to corresponding PHD template entries
        /// 3. Verifies that total hours match between ENA and PHD
        /// </summary>
        /// <param name="enaTimesheet">The ENA timesheet containing hours to update</param>
        /// <exception cref="Exception">Thrown if:
        ///     - Any ENA task is not found in the PHD template
        ///     - Total hours mismatch between ENA and PHD after update
        /// </exception>
        public void Update(EnaTimesheet enaTimesheet)
        {
            CheckTasks(enaTimesheet);
            var enaEffort = enaTimesheet.TotalHoursByClientTaskDay(); // Dictionary<string, Dictionary<int, double>>
            foreach (var phdEntry in entries) // PhdTemplateEntry phdEntry
            {
                if (enaEffort.TryGetValue(phdEntry.ClientHashTask(), out var enaEntry))
                {
                    // Dictionary<int, double> enaEntry contains (day -> hours) for one client task
                    phdEntry.SetEffort(enaEntry);
                }
            }
            double enaTotalHours = enaTimesheet.TotalHours();
            double phdTotalHours = TotalHours();
            if (enaTotalHours != phdTotalHours)
            {
                PrintEntries();
                throw new Exception($"Total hours mismatch: ENA={enaTotalHours} PHD={phdTotalHours}");
            }
            UpdateExcelPackage();
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

        /// <summary>
        /// Updates the ExcelPackage with data from PhdTemplateEntry objects.
        /// </summary>
        public void UpdateExcelPackage()
        {
            var worksheet = _excelPackage.Workbook.Worksheets[0];
            
            foreach (var entry in entries)
            {
                var row = worksheet.Row(entry.RowNum + 1); // +1 because Excel is 1-based
                if (row == null)
                {
                    _logger.LogWarning($"Row {entry.RowNum + 1} not found in worksheet");
                    continue;
                }

                // Clear existing effort cells
                EraseEffort(worksheet, entry.RowNum + 1, row);

                // Set new effort values
                foreach (var dayEffort in entry.GetEffort())
                {
                    int day = dayEffort.Key;
                    double effort = dayEffort.Value;
                    var cell = worksheet.Cells[row.Row, colOffset + day];
                    cell.Value = effort;
                }
            }

            // Recalculate formulas
            _excelPackage.Workbook.Calculate();
        }

        private void EraseEffort(ExcelWorksheet worksheet, int rowNum, ExcelRow row)
        {
            // Clear all cells from colOffset to the end of the row
            for (int col = 1 + colOffset; col <= 31 + colOffset; col++) // Clear up to 31 days
            {
                var cell = worksheet.Cells[row.Row, col];
                cell.Value = null;
            }
        }
    }
}