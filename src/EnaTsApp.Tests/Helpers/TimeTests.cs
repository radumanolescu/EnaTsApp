using System;
using Xunit;
using Ena.Timesheet.Util;

namespace Ena.Timesheet.Tests.Helpers
{
    public class TimeTests
    {
        [Fact]
        public void GetLastDayOfMonth_ValidYearMonths()
        {
            // February (non-leap year)
            Assert.Equal(28, Time.GetLastDayOfMonth("202502"));
            
            // February (leap year)
            Assert.Equal(29, Time.GetLastDayOfMonth("202402"));
            
            // April
            Assert.Equal(30, Time.GetLastDayOfMonth("202504"));
            
            // May
            Assert.Equal(31, Time.GetLastDayOfMonth("202505"));
        }

        [Fact]
        public void GetLastDayOfMonth_InvalidYearMonths()
        {
            // Null input
            Assert.Throws<ArgumentException>(() => Time.GetLastDayOfMonth(null!));
            
            // Empty string
            Assert.Throws<ArgumentException>(() => Time.GetLastDayOfMonth(""));
            
            // Too short
            Assert.Throws<ArgumentException>(() => Time.GetLastDayOfMonth("20250"));
            
            // Too long
            Assert.Throws<ArgumentException>(() => Time.GetLastDayOfMonth("2025040"));
            
            // Invalid year
            Assert.Throws<ArgumentException>(() => Time.GetLastDayOfMonth("abcd04"));
            
            // Invalid month
            Assert.Throws<System.ArgumentOutOfRangeException>(() => Time.GetLastDayOfMonth("202513"));
        }
    }
}
