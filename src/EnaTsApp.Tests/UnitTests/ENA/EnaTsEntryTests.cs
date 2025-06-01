using System;
using Xunit;
using Com.Ena.Timesheet.Ena;
using System.Collections.Generic;

namespace Com.Ena.Timesheet.Tests.UnitTests.ENA
{
    public class EnaTsEntryTests
    {
        [Fact]
        public void CanParseValidEntry()
        {
            // Arrange
            var selectedDate = new DateTime(2025, 4, 1); // April 2025
            var row = new List<string>
            {
                "2305#CD - Pricing/Filing Set Development", // Client#Task
                "7", // Day of month
                "0.458333333", // Start time (11:00 AM)
                "0.479166667", // End time (11:30 AM)
                "0.50", // Hours
                "review pricing" // Description
            };

            // Act
            // Act - Parse the row data
            var entry = new EnaTsEntry(1, selectedDate, row);
            Assert.NotNull(entry);
            Assert.Null(entry.Error); // Verify no parsing errors

            // Assert
            Assert.NotNull(entry);
            // Assert that properties are parsed correctly
            Assert.Equal("2305", entry.ProjectId);
            Assert.Equal("CD - Pricing/Filing Set Development", entry.Activity);
            Assert.Equal(7, entry.Day);
            // Verify time conversion from Excel time values
            // 0.458333333 is 11:00 AM -> 11 hours
            // 0.479166667 is 11:30 AM -> 11.5 hours
            Assert.Equal(11.0, entry.Start.Value.TotalHours);
            Assert.Equal(11.5, entry.End.Value.TotalHours);
            Assert.Equal(0.50f, entry.Hours);
            Assert.Equal("review pricing", entry.Description);
            Assert.Equal(selectedDate, entry.Month);
            Assert.Equal(1, entry.LineId);
        }
    }
}
