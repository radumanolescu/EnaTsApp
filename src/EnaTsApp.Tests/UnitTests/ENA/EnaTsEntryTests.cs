using System;
using Xunit;
using Com.Ena.Timesheet.Ena;
using System.Collections.Generic;
using Xunit.Abstractions;

namespace Com.Ena.Timesheet.Tests.UnitTests.ENA
{
    public class EnaTsEntryTests
    {
        private readonly ITestOutputHelper _output;

        public EnaTsEntryTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void CanParseValidEntry()
        {
            // Arrange
            var selectedDate = new DateTime(2025, 4, 1); // April 2025
            var start = 0.458333333; // Start time (11:00 AM)
            var end = 0.479166667; // End time (11:30 AM)
            var row = new List<string>
            {
                "2305#CD - Pricing/Filing Set Development", // Client#Task
                "7", // Day of month
                start.ToString(), // Start time (11:00 AM)
                end.ToString(), // End time (11:30 AM)
                "0.50", // Hours
                "review pricing" // Description
            };

            // Act - Parse the row data
            var entry = new EnaTsEntry(1, selectedDate, row);
            Assert.NotNull(entry);
            Assert.Equal("", entry.Error); // Verify no parsing errors
            _output.WriteLine($"---------- Error:       [{entry.Error}]");
            _output.WriteLine($"---------- ProjectId:   [{entry.ProjectId}]");
            _output.WriteLine($"---------- Activity:    [{entry.Activity}]");
            _output.WriteLine($"---------- Day:         [{entry.Day}]");
            _output.WriteLine($"---------- Start:       [{entry.Start?.TotalHours}]");
            _output.WriteLine($"---------- End:         [{entry.End?.TotalHours}]");
            _output.WriteLine($"---------- Hours:       [{entry.Hours}]");
            _output.WriteLine($"---------- Description: [{entry.Description}]");
            _output.WriteLine($"---------- Month:       [{entry.Month}]");
            _output.WriteLine($"---------- LineId:      [{entry.LineId}]");

            // Assert
            Assert.NotNull(entry);
            // Assert that properties are parsed correctly
            Assert.Equal("2305", entry.ProjectId);
            Assert.Equal("CD - Pricing/Filing Set Development", entry.Activity);
            Assert.Equal(7, entry.Day);
            // Verify time conversion from Excel time values
            // 0.458333333 is 11:00 AM -> 11 hours
            // 0.479166667 is 11:30 AM -> 11.5 hours
            Assert.Equal(24*start, entry.Start?.TotalHours);
            Assert.Equal(24*end, entry.End?.TotalHours);
            Assert.Equal(0.50f, entry.Hours);
            Assert.Equal("review pricing", entry.Description);
            Assert.Equal(selectedDate, entry.Month);
            Assert.Equal(1, entry.LineId);
        }
    }
}
