using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;
using Com.Ena.Timesheet.Ena;
using EnaTsApp.Tests.TestHelpers;
using Ena.Timesheet.Tests;
using Com.Ena.Timesheet.Xl;
using OfficeOpenXml;

namespace Com.Ena.Timesheet.Tests.IntegrationTests.Ena
{
    public class EnaTimesheetIntegrationTests
    {
        public EnaTimesheetIntegrationTests()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        }

        [Fact]
        public void CanParseEnaTimesheet()
        {

            // Arrange
            var filePath = TestFileHelper.GetTestFilePath("PHD 04 - April 2025R.xlsx");
            var selectedDate = new DateTime(2025, 4, 1); // April 2025
            
            try
            {
                // Act
                var parser = new ExcelParser();
                var rows = parser.ParseExcelFile(filePath);
                Assert.NotNull(rows);

                var logger = TestLogger.CreateLogger<EnaTimesheet>();
                var entryLogger = TestLogger.CreateLogger<EnaTsEntry>();
                var enaTimesheet = new EnaTimesheet(selectedDate.ToString("yyyyMM"), rows, filePath, filePath);
                
                // Assert
                Assert.NotNull(enaTimesheet);
                Assert.NotEmpty(enaTimesheet.GetEntries());
                
                // Verify entries are parsed correctly
                var entries = enaTimesheet.GetEntries();
                Assert.True(entries.Any());
                
                // Check first entry
                var firstEntry = entries.First();
                Assert.NotNull(firstEntry);
                Assert.Equal(1, firstEntry.LineId); // First row should have lineId 1
                Assert.Equal(selectedDate, firstEntry.Month);
                
                // Write entries to file for verification
                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(directory))
                {
                    throw new InvalidOperationException("Test file path is invalid");
                }

                var outputFile = Path.Combine(directory, "ParseEnaTimesheet.txt");
                using (var writer = new StreamWriter(outputFile))
                {
                    foreach (var entry in entries)
                    {
                        writer.WriteLine($"{entry.ToString()}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
        }

    }
}
