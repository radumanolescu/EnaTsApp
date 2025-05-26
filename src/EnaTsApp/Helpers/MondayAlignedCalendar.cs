using System;
using System.Collections.Generic;

namespace Ena.Timesheet.Util
{
    /// <summary>
    /// A calendar that is aligned to Mondays, and has a week number for each day.
    /// Weeks are defined as Monday through Sunday.
    /// Week numbers are 1-based.
    /// </summary>
    public class MondayAlignedCalendar
    {
        private readonly DateTime anchorDate;
        private readonly Dictionary<DateTime, int> weekOfMonth = new Dictionary<DateTime, int>();

        /// <summary>
        /// Constructs a Monday-aligned calendar for the month of the anchor date.
        /// </summary>
        /// <param name="anchorDate">A date within the month to anchor the calendar.</param>
        public MondayAlignedCalendar(DateTime anchorDate)
        {
            this.anchorDate = anchorDate.Date;
            DateTime date = new DateTime(anchorDate.Year, anchorDate.Month, 1);
            int week = 1;
            while (date.Month == anchorDate.Month)
            {
                weekOfMonth[date] = week;
                date = date.AddDays(1);
                if ((int)date.DayOfWeek == 1) // Monday (DayOfWeek: Sunday=0, Monday=1, ...)
                {
                    week++;
                }
            }
        }

        /// <summary>
        /// Gets the week number (1-based) for the given date in the month.
        /// </summary>
        public int? GetWeekOfMonth(DateTime date)
        {
            weekOfMonth.TryGetValue(date.Date, out int week);
            return week == 0 ? (int?)null : week;
        }
    }
}