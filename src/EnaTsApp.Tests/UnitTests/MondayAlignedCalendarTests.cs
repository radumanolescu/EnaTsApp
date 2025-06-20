using System;
using Ena.Timesheet.Util;
using Xunit;

namespace EnaTsApp.Tests.UnitTests
{
    public class MondayAlignedCalendarTests
    {
        [Fact]
        public void TestWeekNumbersForJanuary2024()
        {
            // January 2024 has 5 Mondays (1st, 8th, 15th, 22nd, 29th)
            var calendar = new MondayAlignedCalendar(new DateTime(2024, 1, 15));

            // First week (1st-7th)
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2024, 1, 1)));
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2024, 1, 7)));

            // Second week (8th-14th)
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2024, 1, 8)));
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2024, 1, 14)));

            // Third week (15th-21st)
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2024, 1, 15)));
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2024, 1, 21)));

            // Fourth week (22nd-28th)
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2024, 1, 22)));
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2024, 1, 28)));

            // Fifth week (29th-31st)
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2024, 1, 29)));
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2024, 1, 31)));
        }

        [Fact]
        public void TestWeekNumbersForFebruary2024()
        {
            // February 2024 has 5 Mondays (5th, 12th, 19th, 26th)
            var calendar = new MondayAlignedCalendar(new DateTime(2024, 2, 15));

            // First week (1st-4th)
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2024, 2, 1)));
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2024, 2, 4)));

            // Second week (5th-11th)
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2024, 2, 5)));
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2024, 2, 11)));

            // Third week (12th-18th)
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2024, 2, 12)));
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2024, 2, 18)));

            // Fourth week (19th-25th)
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2024, 2, 19)));
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2024, 2, 25)));

            // Fifth week (26th-29th)
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2024, 2, 26)));
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2024, 2, 29)));
        }

        [Fact]
        public void TestWeekNumbersForApril2025()
        {
            // April 1, 2025 is a Tuesday
            var calendar = new MondayAlignedCalendar(new DateTime(2025, 4, 15));

            // First week (1st-6th)
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2025, 4, 1)));
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2025, 4, 6)));

            // Second week (7th-13th)
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2025, 4, 7)));
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2025, 4, 13)));

            // Third week (14th-20th)
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2025, 4, 14)));
            Assert.Equal(3, calendar.GetWeekOfMonth(new DateTime(2025, 4, 20)));

            // Fourth week (21st-27th)
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2025, 4, 21)));
            Assert.Equal(4, calendar.GetWeekOfMonth(new DateTime(2025, 4, 27)));

            // Fifth week (28th-30th)
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2025, 4, 28)));
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2025, 4, 30)));

            // Edge cases
            Assert.Null(calendar.GetWeekOfMonth(new DateTime(2025, 3, 31))); // Not in April
            Assert.Null(calendar.GetWeekOfMonth(new DateTime(2025, 5, 1)));  // Not in April
        }

        [Fact]
        public void TestWeekNumbersForJanuary2025()
        {
            var calendar = new MondayAlignedCalendar(new DateTime(2025, 1, 1));
            
            Assert.Equal(1, calendar.GetWeekOfMonth(new DateTime(2025, 1, 1)));
            Assert.Equal(2, calendar.GetWeekOfMonth(new DateTime(2025, 1, 6)));
            Assert.Equal(5, calendar.GetWeekOfMonth(new DateTime(2025, 1, 31)));
        }

        [Fact]
        public void TestWeekNumbersForEdgeCases()
        {
            // Test with dates around the month boundaries
            var calendar = new MondayAlignedCalendar(new DateTime(2024, 1, 15));
            
            // Test dates before the month
            Assert.Null(calendar.GetWeekOfMonth(new DateTime(2023, 12, 31)));
            
            // Test dates after the month
            Assert.Null(calendar.GetWeekOfMonth(new DateTime(2024, 2, 1)));
        }
    }
}

// dotnet test src/EnaTsApp.Tests/EnaTsApp.Tests.csproj --filter "FullyQualifiedName~MondayAlignedCalendarTests" --verbosity normal --logger "console;verbosity=detailed"
