using System;
using System.Globalization;

namespace Ena.Timesheet.Util
{
    public static class Time
    {
        public static readonly string IsoLocalDateFmt = "yyyy-MM-dd";
        public static readonly string YyyyMMFmt = "yyyyMM";
        public static readonly string MMYYFmt = "MMMM yyyy";

        public static float? HoursBetween(TimeSpan? start, TimeSpan? end)
        {
            if (start == null || end == null) return null;
            var minutes = (end.Value - start.Value).TotalMinutes;
            return (float)(minutes / 60.0);
        }

        public static int WeekOfMonth(DateTime date)
        {
            // .NET's CalendarWeekRule.FirstDay and DayOfWeek.Monday for Monday-aligned weeks
            var firstDay = new DateTime(date.Year, date.Month, 1);
            int week = 1;
            DateTime d = firstDay;
            while (d <= date)
            {
                if (d.DayOfWeek == DayOfWeek.Monday && d != firstDay)
                    week++;
                d = d.AddDays(1);
            }
            return week;
        }

        /// <summary>
        /// Converts a date string in yyyy-MM-dd format to yyyyMM.
        /// </summary>
        public static string GetYearMonth(string dateStr)
        {
            var tsMonth = DateTime.ParseExact(dateStr, IsoLocalDateFmt, CultureInfo.InvariantCulture);
            return tsMonth.ToString(YyyyMMFmt, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Converts a date string in yyyy-MM-dd format to MMMM yyyy.
        /// </summary>
        public static string GetMonthYear(string dateStr)
        {
            var tsMonth = DateTime.ParseExact(dateStr, IsoLocalDateFmt, CultureInfo.InvariantCulture);
            return tsMonth.ToString(MMYYFmt, CultureInfo.InvariantCulture);
        }
    }
}