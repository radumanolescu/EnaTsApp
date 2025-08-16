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
using F23.StringSimilarity;

namespace Com.Ena.Timesheet.Ena
{
    public class EnaTimesheet : ExcelMapped
    {
        private const int SHEET_INDEX = 0;
        private const int ERROR_COLUMN = 7;

        private readonly DateTime timesheetMonth;
        public DateTime TimesheetMonth => timesheetMonth;

        private readonly List<EnaTsEntry> enaTsEntries = new List<EnaTsEntry>();
        public IReadOnlyList<EnaTsEntry> EnaTsEntries => enaTsEntries;

        private readonly List<EnaTsProjectEntry> projectEntries = new List<EnaTsProjectEntry>();
        public IReadOnlyList<EnaTsProjectEntry> ProjectEntries => projectEntries;

        private readonly ILogger<EnaTimesheet> _logger;

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
            this.timesheetMonth = DateTime.ParseExact(yyyyMM, "yyyyMM", CultureInfo.InvariantCulture);
            populateFromData(yyyyMM, timesheetData);
        }

        private void populateFromData(string yyyyMM, List<List<string>> timesheetData)
        {
            _logger.LogInformation($"Populating timesheet entries for {yyyyMM}");

            for (int i = 0; i < timesheetData.Count; i++)
            {
                var row = timesheetData[i];
                // Ignore any row that is empty or contains only whitespace
                if (row.Count == 0 || row.All(cell => string.IsNullOrWhiteSpace(cell)))
                {
                    continue;
                }
                var entry = new EnaTsEntry(i, timesheetMonth, row);
                enaTsEntries.Add(entry);
                _logger.LogInformation($"Added entry for row {i}: {entry.ProjectActivity()}");
            }
            ParseSortReindexEntries(enaTsEntries);
        }

        private void ParseSortReindexEntries(List<EnaTsEntry> inputEntries)
        {
            // Sort and reindex entries in-place
            SortByDayProjectId(inputEntries);
            ReindexEntries(inputEntries);
            this.projectEntries.Clear();
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

        public string GetEntriesWithTotalsAsHtml()
        {
            // Since Cottle doesn't seem to support list iteration, we need to join the rows manually
            var entriesWithTotals = GetEntriesWithTotals();
            return string.Join("\n", entriesWithTotals.Select(e => e.GetHtmlRow()));
        }

        public List<EnaTsEntry> GetEntries() => enaTsEntries;

        public List<EnaTsProjectEntry> GetProjectEntries() => projectEntries;

        public string GetProjectEntriesAsHtml()
        {
            var entriesWithTotals = GetProjectEntries();
            return string.Join("\n", entriesWithTotals.Select(e => e.GetHtmlRow()));
        }

        /// <summary>
        /// Updates the Excel file by writing validation error messages for each entry that has an error.
        /// The error messages are written to the error column (ERROR_COLUMN) of the worksheet.
        /// The updated Excel object is saved to the output directory as a new Excel file.
        /// </summary>
        private void UpdateExcelFile()
        {
            var worksheet = _excelPackage.Workbook.Worksheets[SHEET_INDEX];
            // Write the entry error message (if any) to the error column
            foreach (var entry in enaTsEntries)
            {
                if (string.IsNullOrEmpty(entry.Error))
                    continue;
                var row = worksheet.Row(entry.LineId + 1); // +1 because Excel is 1-based
                if (row == null)
                    continue;
                var cell = worksheet.Cells[row.Row, ERROR_COLUMN];
                cell.Value = entry.Error;
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
            }
            SaveAs();
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

        /// <summary>
        /// Updates the activity for all entries that match the invalid activity to the valid activity.
        /// After updating, it saves the timesheet using SaveAs().
        /// </summary>
        /// <param name="invalidActivity">The activity to be replaced</param>
        /// <param name="validActivity">The new activity to set</param>
        public void UpdateActivity(string invalidActivity, string validActivity)
        {
            if (string.IsNullOrEmpty(invalidActivity) || string.IsNullOrEmpty(validActivity))
                return;

            bool updated = false;
            foreach (var entry in enaTsEntries)
            {
                if (string.Equals(entry.Activity, invalidActivity, StringComparison.OrdinalIgnoreCase))
                {
                    entry.Activity = validActivity;
                    updated = true;
                }
            }

            if (updated)
            {
                SaveAs();
            }
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

        /// <summary>
        /// Validates all client tasks in the timesheet against a set of allowed client tasks.
        /// For each entry that has an invalid project activity:
        /// 1. Marks the entry as invalid
        /// 2. Suggests the closest matching client task from the allowed set
        /// 3. Updates the Excel file with validation errors if any are found
        /// </summary>
        /// <param name="clientTaskSet">A set of valid client task strings that entries must match.</param>
        /// <returns>True if all entries have valid project activities, false otherwise.</returns>
        public bool IsValidAllClientTasks(HashSet<string> clientTaskSet)
        {
            bool valid = true;
            foreach (var entry in enaTsEntries)
            {
                if (!string.IsNullOrEmpty(entry.Error))
                {
                    valid = false;
                }
                if (!clientTaskSet.Contains(entry.ProjectActivity()))
                {
                    valid = false;
                    string err0 = entry.Error;
                    string suggested = BestMatch(entry.ProjectActivity(), clientTaskSet);
                    string err1 = $"Invalid project#activity. Did you mean '{suggested}'?";
                    _logger.LogInformation($"NotInClientTaskSet [{entry.LineId}] {entry.ProjectActivity()} -> {suggested}");
                    _logger.LogInformation(err0 + err1);
                    entry.SetError(err0 + err1);
                }
            }
            if (!valid)
            {
                UpdateExcelFile();
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

        /// <summary>
        /// Finds the best matching client task from a set of tasks using Jaro-Winkler string similarity.
        /// This method is used to suggest corrections for invalid project activities in timesheet entries.
        /// </summary>
        /// <param name="projectActivity">The project activity string to find a match for.</param>
        /// <param name="clientTaskSet">A set of valid client task strings to search through.</param>
        /// <returns>The best matching client task string from the set, or an empty string if no valid match is found.</returns>
        /// <remarks>
        /// Uses the Jaro-Winkler algorithm to calculate string similarity, which is particularly
        /// good at matching strings with common prefixes and small transpositions.
        /// </remarks>
        public static string BestMatch(string projectActivity, HashSet<string> clientTaskSet)
        {
            if (string.IsNullOrEmpty(projectActivity) || clientTaskSet == null || clientTaskSet.Count == 0)
                return "";

            string bestMatch = "";
            double highestScore = 0;

            foreach (string candidate in clientTaskSet)
            {
                double score = new JaroWinkler().Similarity(projectActivity, candidate);
                if (score > highestScore)
                {
                    highestScore = score;
                    bestMatch = candidate;
                }
            }

            return bestMatch;
        }

        /// <summary>
        /// Adds 'R' as the last character before the extension in an Excel filename.
        /// For example, "PHD 04 - April 2025.xlsx" becomes "PHD 04 - April 2025R.xlsx"
        /// </summary>
        /// <param name="filename">The input filename to modify. Must be a non-empty string with an extension.</param>
        public static string AddRevisionToFilename(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty", nameof(filename));

            string extension = Path.GetExtension(filename);
            if (string.IsNullOrEmpty(extension))
                throw new ArgumentException("Filename must have an extension", nameof(filename));

            string baseName = Path.GetFileNameWithoutExtension(filename);
            return $"{baseName}R{extension}";
        }

        // Helper class for weekly summary
        private class WeeklySummary
        {
            public float TotalHours = 0.0f;
            public float MaxEntryId = 0.0f;
            public float TotalCharge = 0.0f;
        }
    }
}