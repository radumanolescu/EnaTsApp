using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Collections.ObjectModel;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using Com.Ena.Timesheet;
using NPOI.XSSF.UserModel;

namespace Com.Ena.Timesheet.Ena
{
    public class EnaTimesheet : ExcelMapped
    {
        private const int SHEET_INDEX = 0;

        private readonly DateTime timesheetMonth;
        private readonly List<EnaTsEntry> enaTsEntries = new List<EnaTsEntry>();
        private List<EnaTsProjectEntry> projectEntries = new List<EnaTsProjectEntry>();
        private byte[]? xlsxBytes;
        private readonly ILogger<EnaTimesheet> _logger;
        private readonly ILogger<EnaTsEntry> _entryLogger;

        private static readonly IServiceProvider _serviceProvider = new ServiceCollection()
            .AddLogging(builder => builder.AddConsole())
            .BuildServiceProvider();

        private static ILogger<T> GetLogger<T>()
        {
            return _serviceProvider.GetRequiredService<ILogger<T>>();
        }

        public EnaTimesheet(string yyyyMM, List<List<string>> timesheetData, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<EnaTimesheet>();
            _entryLogger = GetLogger<EnaTsEntry>();
            this.timesheetMonth = DateTime.ParseExact(yyyyMM, "yyyyMM", CultureInfo.InvariantCulture);
            populateFromData(yyyyMM, timesheetData);
        }

        private void populateFromData(string yyyyMM, List<List<string>> timesheetData)
        {
            _logger.LogInformation($"Populating timesheet entries for {yyyyMM}");
            
            // Skip header row
            for (int i = 1; i < timesheetData.Count; i++)
            {
                var row = timesheetData[i];
                if (row.Count > 0)
                {
                    var entry = new EnaTsEntry(i, timesheetMonth, row);
                    enaTsEntries.Add(entry);
                    _logger.LogInformation($"Added entry for row {i}: {entry.ProjectActivity()}");
                }
            }
        }

        private byte[] CreateXlsxFromData(List<List<string>> timesheetData)
        {
            using var ms = new MemoryStream();
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet();

            // Add headers
            var headerRow = sheet.CreateRow(0);
            foreach (var header in timesheetData[0])
            {
                headerRow.CreateCell(headerRow.LastCellNum).SetCellValue(header);
            }

            // Add data rows
            for (int i = 1; i < timesheetData.Count; i++)
            {
                var row = sheet.CreateRow(i);
                foreach (var cellValue in timesheetData[i])
                {
                    row.CreateCell(row.LastCellNum).SetCellValue(cellValue);
                }
            }

            workbook.Write(ms);
            return ms.ToArray();
        }

        public EnaTimesheet(DateTime timesheetMonth, Stream inputStream, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<EnaTimesheet>();
            _entryLogger = GetLogger<EnaTsEntry>();
            this.timesheetMonth = timesheetMonth;
            this.xlsxBytes = GetBytes(inputStream);
            using (var ms = new MemoryStream(this.xlsxBytes))
            {
                ParseSortReindex(ms);
            }
        }

        public EnaTimesheet(DateTime timesheetMonth, FileInfo enaTimesheetFile, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<EnaTimesheet>();
            _entryLogger = GetLogger<EnaTsEntry>();
            this.timesheetMonth = timesheetMonth;
            using (var fs = enaTimesheetFile.OpenRead())
            {
                this.xlsxBytes = GetBytes(fs);
            }
            using (var ms = new MemoryStream(this.xlsxBytes))
            {
                ParseSortReindex(ms);
            }
        }

        public EnaTimesheet(DateTime timesheetMonth, byte[] fileBytes, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<EnaTimesheet>();
            _entryLogger = GetLogger<EnaTsEntry>();
            this.timesheetMonth = timesheetMonth;
            this.xlsxBytes = fileBytes;
            using (var ms = new MemoryStream(this.xlsxBytes))
            {
                ParseSortReindex(ms);
            }
        }

        public EnaTimesheet(DateTime timesheetMonth, List<List<string>> timesheetData, string inputPath, string outputPath)
            : base(inputPath, outputPath)
        {
            _logger = GetLogger<EnaTimesheet>();
            _entryLogger = GetLogger<EnaTsEntry>();
            this.timesheetMonth = timesheetMonth;
            this.enaTsEntries = new List<EnaTsEntry>();
            xlsxBytes = null;
            _logger.LogInformation($"Creating EnaTimesheet for month {timesheetMonth.ToString("yyyy-MM")}");
            
            if (timesheetData != null)
            {
                for (int i = 0; i < timesheetData.Count; i++)
                {
                    var entry = new EnaTsEntry(i, timesheetMonth, timesheetData[i]);
                    this.enaTsEntries.Add(entry);
                }
            }
        }

        private void ParseSortReindex(Stream inputStream)
        {
            var inputEntries = ParseEntries(inputStream);
            SortByDayProjectId(inputEntries);
            ReindexEntries(inputEntries);
            this.enaTsEntries.AddRange(inputEntries);
            this.projectEntries.AddRange(GetProjectEntries(inputEntries));
        }

        public List<EnaTsEntry> GetEntriesWithTotals()
        {
            var entriesWithTotals = new List<EnaTsEntry>(enaTsEntries);
            var totalEntries = WeeklyTotals(enaTsEntries);
            entriesWithTotals.AddRange(totalEntries);
            SortByEntryId(entriesWithTotals);
            return entriesWithTotals;
        }

        public List<EnaTsEntry> GetEntries() => enaTsEntries;

        public List<EnaTsProjectEntry> GetProjectEntries() => projectEntries;

        public byte[] GetXlsxBytes() => xlsxBytes;
        public void SetXlsxBytes(byte[] bytes) => xlsxBytes = bytes;

        private List<EnaTsEntry> ParseEntries(Stream inputStream)
        {
            var enaEntries = new List<EnaTsEntry>();
            // You need to use a library like EPPlus or ClosedXML to read Excel files in C#
            // This is a placeholder for actual Excel parsing logic
            // For each row in the sheet, create an EnaTsEntry and add to enaEntries
            // Example:
            // using (var package = new ExcelPackage(inputStream)) { ... }
            // For now, this is left as a stub.
            throw new NotImplementedException("Excel parsing not implemented. Use EPPlus or ClosedXML.");
        }

        private void UpdateBytes(/*Workbook workbook*/ float notUsed = 0.0f)
        {
            // Use EPPlus or ClosedXML to write the workbook to a MemoryStream and update xlsxBytes
            throw new NotImplementedException("Excel writing not implemented. Use EPPlus or ClosedXML.");
        }

        private void UpdateBytes()
        {
            // Use EPPlus or ClosedXML to update error cells and write to xlsxBytes
            throw new NotImplementedException("Excel writing not implemented. Use EPPlus or ClosedXML.");
        }

        private void ReindexEntries(List<EnaTsEntry> entries)
        {
            int lineId = 0;
            foreach (var entry in entries)
            {
                entry.SetEntryId((float)lineId++);
            }
        }

        private void SortByDayProjectId(List<EnaTsEntry> entries)
        {
            entries.Sort((a, b) => a.SortKey().CompareTo(b.SortKey()));
        }

        private void SortByEntryId(List<EnaTsEntry> entries)
        {
            entries.Sort((a, b) => Nullable.Compare(a.EntryId, b.EntryId));
        }

        private List<EnaTsEntry> WeeklyTotals(List<EnaTsEntry> entries)
        {
            var totalsByWeek = new List<EnaTsEntry>();
            var weeklySummaries = GetWeeklySummary(entries);
            float hoursForMonth = 0.0f;
            float chargeForMonth = 0.0f;
            float maxEntryIdForMonth = 0.0f;
            foreach (var entry in weeklySummaries)
            {
                float totalHours = entry.Value.TotalHours;
                float maxEntryId = entry.Value.MaxEntryId;
                float totalCharge = entry.Value.TotalCharge;
                totalsByWeek.Add(new EnaTsWeekTotalEntry(maxEntryId + 0.1f, "Total hours: ", totalHours, "Total for week:", totalCharge));
                hoursForMonth += totalHours;
                chargeForMonth += totalCharge;
                maxEntryIdForMonth = Math.Max(maxEntryIdForMonth, maxEntryId);
                totalsByWeek.Add(new EnaTsWeekBlankEntry(maxEntryId + 0.2f));
                totalsByWeek.Add(new EnaTsWeekBlankEntry(maxEntryId + 0.3f));
            }
            totalsByWeek.Add(new EnaTsWeekTotalEntry(maxEntryIdForMonth + 0.4f, "Monthly hours:", hoursForMonth, "Total consulting fees for month:", chargeForMonth));
            totalsByWeek.Add(new EnaTsWeekBlankEntry(maxEntryIdForMonth + 0.5f));
            return totalsByWeek;
        }

        private Dictionary<int, WeeklySummary> GetWeeklySummary(List<EnaTsEntry> entries)
        {
            var summaryByWeek = new Dictionary<int, WeeklySummary>();
            foreach (var entry in entries)
            {
                int week = entry.GetWeekOfMonth();
                if (!summaryByWeek.TryGetValue(week, out var summary))
                {
                    summary = new WeeklySummary();
                    summaryByWeek[week] = summary;
                }
                summary.TotalHours += entry.Hours ?? 0.0f;
                summary.MaxEntryId = Math.Max(summary.MaxEntryId, entry.EntryId ?? 0.0f);
                summary.TotalCharge += entry.Charge ?? 0.0f;
            }
            return summaryByWeek;
        }

        public DateTime GetTimesheetMonth() => timesheetMonth;

        private List<EnaTsProjectEntry> GetProjectEntries(List<EnaTsEntry> entries)
        {
            var projectEntries = new List<EnaTsProjectEntry>();
            float totalHours = 0.0f;
            var projectEntryMap = new Dictionary<string, EnaTsProjectEntry>();
            foreach (var entry in entries)
            {
                string projectActivity = entry.ProjectActivity();
                if (!projectEntryMap.TryGetValue(projectActivity, out var projectEntry))
                {
                    projectEntry = new EnaTsProjectEntry(entry.ProjectId, entry.Activity, 0.0f);
                }
                projectEntry.Hours += entry.Hours ?? 0.0f;
                totalHours += entry.Hours ?? 0.0f;
                projectEntryMap[projectActivity] = projectEntry;
            }
            projectEntries.AddRange(projectEntryMap.Values);
            projectEntries.Sort();
            projectEntries.Add(new EnaTsProjectEntry("Total hours:", "", totalHours));
            return projectEntries;
        }

        public double TotalHours()
        {
            return enaTsEntries.Sum(e => e.Hours ?? 0.0f);
        }

        public Dictionary<string, Dictionary<int, double>> TotalHoursByClientTaskDay()
        {
            return enaTsEntries
                .GroupBy(e => e.ProjectActivity())
                .ToDictionary(
                    g => g.Key,
                    g => g.GroupBy(e => e.Day)
                          .ToDictionary(
                              gg => gg.Key ?? 0,
                              gg => gg.Sum(e => e.Hours ?? 0.0)
                          )
                );
        }

        public bool IsValidAllClientTasks(HashSet<string> clientTaskSet)
        {
            bool valid = true;
            foreach (var entry in enaTsEntries)
            {
                if (!clientTaskSet.Contains(entry.ProjectActivity()))
                {
                    valid = false;
                    string err0 = entry.Error;
                    string suggested = BestMatch(entry.ProjectActivity(), clientTaskSet);
                    string err1 = $"Invalid project#activity. Did you mean '{suggested}'?";
                    Console.WriteLine($"NotInClientTaskSet [{entry.LineId}] {entry.ProjectActivity()} -> {suggested}");
                    Console.WriteLine(err0 + err1);
                    entry.SetError(err0 + err1);
                }
            }
            if (!valid)
            {
                UpdateBytes();
            }
            return valid;
        }

        public bool IsValid()
        {
            return enaTsEntries.All(e => string.IsNullOrEmpty(e.Error));
        }

        public int NumValidationErrors()
        {
            return enaTsEntries.Count(e => !string.IsNullOrEmpty(e.Error));
        }

        public static string BestMatch(string projectActivity, HashSet<string> clientTaskSet)
        {
            // Use a string similarity algorithm, e.g., Jaro-Winkler or Levenshtein
            // For now, return the first entry as a stub
            // You can use the F23.StringSimilarity or SimMetrics.Net NuGet packages for real implementation
            return clientTaskSet.FirstOrDefault() ?? "";
        }

        public void WriteFile(FileInfo output)
        {
            File.WriteAllBytes(output.FullName, xlsxBytes);
        }

        // Helper to read all bytes from a stream
        private static byte[] GetBytes(Stream input)
        {
            using (var ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }

        // Helper class for weekly summary
        private class WeeklySummary
        {
            public float TotalHours = 0.0f;
            public float MaxEntryId = 0.0f;
            public float TotalCharge = 0.0f;
        }
    }

    // You need to implement EnaTsEntry, EnaTsProjectEntry, EnaTsWeekTotalEntry, EnaTsWeekBlankEntry for this to work.
}