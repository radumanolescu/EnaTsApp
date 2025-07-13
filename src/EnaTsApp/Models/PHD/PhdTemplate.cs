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
using Ena.Timesheet.Util;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;
using Com.Ena.Timesheet.PHD;

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

            In the PHD Template:
            - the top row (Excel row 1, template entry.RowNum = 0) is the header row.
            - the first column (Excel column 1) is the client column.
            - the second column (Excel column 2) is the task column.
            - the first day of the month (Excel column 3) is in cell C2 (row=2, col=3)
            So when mapping day D of the month, we need a row offset and a column offset.
        */
        public static readonly int dayColOffset = 2; // Column offset for day D into the Excel coordinates

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
        public bool Update(EnaTimesheet enaTimesheet)
        {
            if (!enaTimesheet.IsValidAllClientTasks(ClientTaskSet()))
            {
                throw new Exception($"Found invalid tasks, see timesheet file {enaTimesheet.OutputPath}");
            }
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
            CheckOverallHours();
            return true;
        }

        private void PrintEntries()
        {
            foreach (var entry in entries)
            {
                Console.WriteLine(entry.ClientTaskEffort());
            }
        }

        public void CheckTasks(EnaTimesheet enaTimesheet) // UNUSED
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
            // entry.RowNum is 0-based, so we need to add 1 to get the correct Excel row number
            // entry.RowNum = 0 is the first row of the template, containing the headers
            
            foreach (var entry in entries)
            {
                var row = worksheet.Row(entry.RowNum + 1); // +1 because Excel is 1-based
                if (row == null)
                {
                    _logger.LogWarning($"Row {entry.RowNum + 1} not found in worksheet");
                    continue;
                }

                // Clear existing effort cells, but not on the header row (entry.RowNum = 0)
                if (entry.RowNum > 0)
                {
                    int excelRowNum = entry.RowNum + 1;
                    EraseEffort(worksheet, excelRowNum, row);
                }

                // Set new effort values
                foreach (var dayEffort in entry.GetEffort())
                {
                    int day = dayEffort.Key;
                    double effort = dayEffort.Value;
                    var cell = worksheet.Cells[row.Row, dayColOffset + day];
                    cell.Value = effort;
                }
            }

            // Recalculate formulas
            _excelPackage.Workbook.Calculate();
        }

        /// <summary>
        /// Verifies that the total hours in the Excel file match the total hours in the object model.
        /// Performs two checks:
        /// 1. Compares the SUM row total with the model's total hours
        /// 2. Compares the SUM+1 row total with the model's total hours
        /// This is necessary because sometimes the Excel formulas are not correct in the template,
        /// and they end up excluding certain columns.
        /// </summary>
        /// <exception cref="System.Exception">Thrown if either total hours check fails.</exception>
        public void CheckOverallHours()
        {
            var worksheet = _excelPackage.Workbook.Worksheets[0];
            int columnOfTotals = excelColumnOf("TOTALS", 1); // On the top row, index of "TOTALS"
            if (columnOfTotals == 0)
            {
                throw new PhdException("TOTALS column not found");
            }
            int rowOfSum = excelRowOf("SUM", 1); // In the first column, index of "SUM"
            if (rowOfSum == 0)
            {
                throw new PhdException("SUM row not found");
            }
            double modelTotalHours = TotalHours(); // in the object model
            var sum1 = worksheet.Cells[rowOfSum, columnOfTotals].Value;
            if (sum1 == null)
            {
                throw new PhdException($"SUM cell not found at row {rowOfSum}, column {columnOfTotals}");
            }
            double sumTotalHours = Convert.ToDouble(sum1); // in the Excel file
            if (modelTotalHours != sumTotalHours)
            {
                throw new PhdException($"Total hours mismatch: Model={modelTotalHours} != SUM={sumTotalHours}");
            }
            // The final total for the month is on the next row
            var sum2 = worksheet.Cells[rowOfSum + 1, columnOfTotals].Value;
            if (sum2 == null)
            {
                throw new PhdException($"SUM+1 cell not found at row {rowOfSum + 1}, column {columnOfTotals}");
            }
            sumTotalHours = Convert.ToDouble(sum2); // in the Excel file
            if (modelTotalHours != sumTotalHours)
            {
                throw new PhdException($"Total hours mismatch: Model={modelTotalHours} != SUM+1={sumTotalHours}");
            }
        }   

        private void EraseEffort(ExcelWorksheet worksheet, int rowNum, ExcelRow row)
        {
            // Clear all cells from first day of the month to the end of the month
            int lastDayOfMonth = Time.GetLastDayOfMonth(yearMonth);
            for (int day = 1; day <= lastDayOfMonth; day++) // Clear up to last day of the month
            {
                var cell = worksheet.Cells[row.Row, dayColOffset + day];
                cell.Value = null;
            }
        }

        /// <summary>
        /// Gets the standardized filename for the (filled-in) PHD timesheet Excel file.
        /// The filename format is "PHD ENA Timesheet YYYY-MM.xlsx" where YYYY-MM is the year and month.
        /// </summary>
        /// <param name="yearMonth">The year and month in YYYYMM format (e.g., "202504").</param>
        /// <returns>The formatted PHD timesheet filename.</returns>
        public static string GetTimesheetFileName(string yearMonth)
        {
            // e.g. "PHD ENA Timesheet 2023-03.xlsx"
            return "PHD ENA Timesheet " + yearMonth.Substring(0, 4) + "-" + yearMonth.Substring(4, 2) + ".xlsx";
        }


    }
}