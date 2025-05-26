using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;

namespace Ena.Timesheet.Ena
{
    /// <summary>
    /// Model for a single row in the ENA timesheet
    /// </summary>
    public class EnaTsEntry : IComparable<EnaTsEntry>
    {
        private static readonly NumberFormatInfo decimalFormat = new NumberFormatInfo { NumberDecimalDigits = 2 };
        private static readonly string dateFormat = "MM/dd/yy";
        protected static readonly float HourlyRate = 60.0f;

        private LocalDate month;
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
        private LocalDate date;
        protected readonly int lineId;
        private MondayAlignedCalendar calendar;

        public EnaTsEntry()
        {
            this.lineId = -1;
        }

        // You will need to implement or replace MondayAlignedCalendar, LocalDate, and Text.unquote for C#
        public EnaTsEntry(int lineId, LocalDate month, IList<object> row)
        {
            this.lineId = lineId;
            this.entryId = (float)lineId;
            this.month = month;
            this.calendar = new MondayAlignedCalendar(month);

            var err = new StringBuilder();
            int cellId = 0;
            int validCells = 0;
            foreach (var cell in row)
            {
                switch (cellId)
                {
                    case 0: // projectId#activity
                        try
                        {
                            string projectActivity = cell?.ToString() ?? "";
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
                            this.date = month.WithDayOfMonth(day.Value);
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
                            this.description = cell?.ToString() ?? "";
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
                            err.Append(cell?.ToString());
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

        public LocalDate Month => month;
        public string ProjectId => projectId;
        public string Activity => activity;
        public int? Day => day;
        public TimeSpan? Start => start;
        public TimeSpan? End => end;
        public float? Hours => hours;
        public string Description => description;
        public float? Charge => charge;
        public string Error => error;
        public float? EntryId 
        { 
            get => entryId;
            protected set => entryId = value;
        }
        public int ValidCells => validCells;
        public int LineId => lineId;

        public string GetRate()
        {
            return $"${HourlyRate.ToString("0.##", decimalFormat)}/hr";
        }

        public string GetCharge()
        {
            return charge.HasValue ? $"${charge.Value:F2}" : "$0.00";
        }

        public string GetDate()
        {
            return date.ToString(dateFormat, CultureInfo.InvariantCulture);
        }

        public int GetWeekOfMonth()
        {
            return calendar.GetWeekOfMonth(date);
        }

        public string SortKey()
        {
            return $"{day.GetValueOrDefault():D5}{projectId}";
        }

        public string ProjectActivity()
        {
            return Unquote(projectId) + "#" + Unquote(activity);
        }

        public string FormattedHours()
        {
            return hours.HasValue ? hours.Value.ToString("0.##", decimalFormat) : "0";
        }

        public override bool Equals(object obj)
        {
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

        public int CompareTo(EnaTsEntry that)
        {
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
        private static TimeSpan? ParseTime(object cell)
        {
            if (cell == null) return null;
            if (cell is TimeSpan ts) return ts;
            if (TimeSpan.TryParse(cell.ToString(), out var result)) return result;
            if (DateTime.TryParse(cell.ToString(), out var dt)) return dt.TimeOfDay;
            return null;
        }

        private static string Unquote(string s)
        {
            return string.IsNullOrEmpty(s) ? "" : s.Replace("\"", "");
        }

        // Placeholder for LocalDate and MondayAlignedCalendar
        // You need to implement or use a suitable library for LocalDate and MondayAlignedCalendar in C#
        // For now, you can use DateTime for LocalDate and stub MondayAlignedCalendar as needed
    }

    // Placeholder for LocalDate (replace with NodaTime or your own implementation)
    public struct LocalDate
    {
        private DateTime date;
        public LocalDate(DateTime dt) { date = dt.Date; }
        public LocalDate WithDayOfMonth(int day) => new LocalDate(new DateTime(date.Year, date.Month, day));
        public override string ToString() => date.ToString("yyyy-MM-dd");
        public string ToString(string format, IFormatProvider provider) => date.ToString(format, provider);
        public override bool Equals(object obj) => obj is LocalDate ld && date.Equals(ld.date);
        public override int GetHashCode() => date.GetHashCode();
    }

    // Placeholder for MondayAlignedCalendar
    public class MondayAlignedCalendar
    {
        private LocalDate month;
        public MondayAlignedCalendar(LocalDate month) { this.month = month; }
        public int GetWeekOfMonth(LocalDate date)
        {
            // Implement week-of-month logic as needed
            return 1;
        }
    }
}