using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;

namespace Com.Ena.Timesheet.Ena
{
    /// <summary>
    /// Model for a single row in the ENA timesheet
    /// </summary>
    public class EnaTsEntry : IComparable<EnaTsEntry>
    {
        private readonly ILogger<EnaTsEntry> _logger;
        private static readonly NumberFormatInfo decimalFormat = new NumberFormatInfo { NumberDecimalDigits = 2 };
        private static readonly string dateFormat = "MM/dd/yy";
        protected static readonly float HourlyRate = 60.0f;

        private static TimeSpan? ParseTime(object timeValue)
        {
            if (timeValue == null) return null;
            
            try
            {
                // Excel stores times as fractions of a day (1.0 = 24 hours)
                double excelTime = Convert.ToDouble(timeValue);
                
                // Convert fraction of day to hours
                double hours = excelTime * 24;
                
                // Create TimeSpan from hours
                return TimeSpan.FromHours(hours);
            }
            catch
            {
                return null;
            }
        }

        private DateTime month;
        private string projectId = "";
        private string activity = "";
        private int? day;
        private TimeSpan? start;
        private TimeSpan? end;
        protected float? hours;
        private string description = "";
        protected float? charge;
        private string error = "";
        private int validCells;
        protected float? entryId;
        private DateTime date;
        protected readonly int lineId;

        public EnaTsEntry()
        {
            this.lineId = -1;
            _logger = null;
        }

        public EnaTsEntry(int lineId, DateTime month, List<string> row, ILogger<EnaTsEntry> logger)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }
            _logger = logger;
            this.lineId = lineId;
            this.entryId = (float)lineId;
            this.month = month;

            var err = new StringBuilder();
            int cellId = 0;
            int validCells = 0;
            // concatenate all cells into a single string for logging
            var rowString = string.Join(", ", row);
            _logger.LogInformation($"Processing row {rowString}");
            foreach (var cell in row)
            {
                switch (cellId)
                {
                    case 0: // projectId#activity
                        try
                        {
                            string projectActivity = cell;
                            var pa = projectActivity.Split('#');
                            if (pa.Length == 2)
                            {
                                this.projectId = pa[0];
                                this.activity = pa[1];
                                validCells++;
                            }
                            else
                            {
                                this.projectId = "";
                                this.activity = "";
                                err.Append($"ProjectActivity must be in the form 'ProjectID#Activity', but was '{projectActivity}'. ");
                            }
                        }
                        catch
                        {
                            err.Append("ProjectActivity must be in the form 'ProjectID#Activity'. ");
                        }
                        break;
                    case 1: // day of month
                        try
                        {
                            this.day = Convert.ToInt32(cell);
                            this.date = new DateTime(month.Year, month.Month, day.Value);
                            validCells++;
                        }
                        catch
                        {
                            err.Append("Day must be a number. ");
                        }
                        break;
                    case 2: // start time
                        try
                        {
                            this.start = ParseTime(cell);
                            _logger.LogDebug($"Parsed start time {cell} to {this.start}");
                            validCells++;
                        }
                        catch
                        {
                            err.Append("Start time must be a time. ");
                        }
                        break;
                    case 3: // end time
                        try
                        {
                            this.end = ParseTime(cell);
                            _logger.LogDebug($"Parsed end time {cell} to {this.end}");
                            validCells++;
                        }
                        catch
                        {
                            err.Append("End time must be a time. ");
                        }
                        break;
                    case 4: // hours
                        try
                        {
                            this.hours = Convert.ToSingle(cell);
                            validCells++;
                        }
                        catch
                        {
                            err.Append("Hours must be a number. ");
                        }
                        break;
                    case 5: // description
                        try
                        {
                            this.description = cell;
                            if (string.IsNullOrEmpty(this.description))
                            {
                                err.Append("Description must be non-empty. ");
                            }
                            else
                            {
                                validCells++;
                            }
                        }
                        catch
                        {
                            err.Append("Description must be non-empty. ");
                        }
                        break;
                    case 6: // validation error, if previously found
                        try
                        {
                            err.Append(cell);
                        }
                        catch
                        {
                            err.Append("Error while reading the error message from the sheet. ");
                        }
                        break;
                }
                cellId++;
            }
            if (this.hours != null)
            {
                this.charge = hours * HourlyRate;
            }
            this.error = err.ToString();
            this.validCells = validCells;
        }

        public DateTime Month => month;
        public string ProjectId { get; set; } = "";
        public string Activity { get; set; } = "";
        public int? Day { get; set; }
        public TimeSpan? Start { get; set; }
        public TimeSpan? End { get; set; }
        public float? Hours { get; set; }
        public DateTime? Date => date;
        public string Description { get; set; } = "";
        public float? Charge { get; set; }
        public string Error { get; set; } = "";

        public void SetError(string error)
        {
            this.Error = error;
        }
        public float? EntryId 
        { 
            get => entryId;
            protected set => entryId = value;
        }
        public int ValidCells => validCells;
        public int LineId => lineId;

        public virtual string GetDate()
        {
            return DateTime.Now.ToString("yyyy-MM-dd");
        }

        public int GetWeekOfMonth()
        {
            // Get the first day of the month
            DateTime firstDayOfMonth = new DateTime(Month.Year, Month.Month, 1);
            // Get the first Monday of the month
            int firstMonday = (8 - (int)firstDayOfMonth.DayOfWeek + (int)DayOfWeek.Monday) % 7;
            // Calculate the week number
            int weekNumber = (Day.GetValueOrDefault() + firstMonday - 1) / 7;
            return weekNumber + 1;
        }

        public virtual string FormattedHours()
        {
            return Hours.HasValue ? string.Format("{0:F2}", Hours.Value) : "";
        }

        public virtual string GetRate()
        {
            return HourlyRate.ToString("0.##", decimalFormat);
        }

        public virtual string GetCharge()
        {
            return Charge.HasValue ? string.Format("{0:C2}", Charge.Value) : "";
        }

        public string SortKey()
        {
            return $"{day.GetValueOrDefault():D5}{projectId}";
        }

        public string ProjectActivity()
        {
            return Unquote(projectId) + "#" + Unquote(activity);
        }

        public override bool Equals(object? obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }
            if (obj is EnaTsEntry that)
            {
                return month.Equals(that.month) && day.Equals(that.day) && start.Equals(that.start);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(month, day, start);
        }

        public int CompareTo(EnaTsEntry? that)
        {
            if (that == null) throw new ArgumentNullException(nameof(that));
            return this.SortKey().CompareTo(that.SortKey());
        }

        public override string ToString()
        {
            return $"{entryId}, {projectId},{activity}, {GetDate()}, {hours}, {start}, {end}, {HourlyRate}, {description}, {charge}";
        }

        public void SetEntryId(float? newId)
        {
            this.EntryId = newId;
        }

        // Helper methods for C# version


        private static string Unquote(string s)
        {
            return string.IsNullOrEmpty(s) ? "" : s.Replace("\"", "");
        }

        // Placeholder for LocalDate
        // You need to implement or use a suitable library for LocalDate in C#
        // For now, you can use DateTime for LocalDate as needed
    }

    // Placeholder for LocalDate (replace with NodaTime or your own implementation)
    public struct LocalDate
    {
        private DateTime date;
        public LocalDate(DateTime dt) { date = dt.Date; }
        public LocalDate WithDayOfMonth(int day) => new LocalDate(new DateTime(date.Year, date.Month, day));
        public override string ToString() => date.ToString("yyyy-MM-dd");
        public string ToString(string format, IFormatProvider provider) => date.ToString(format, provider);
        public override bool Equals(object obj) => obj != null && obj is LocalDate ld && date.Equals(ld.date);
        public override int GetHashCode() => date.GetHashCode();
    }

}