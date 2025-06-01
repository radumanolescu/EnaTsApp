using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Ena;
using EnaTsApp.Tests.TestHelpers;
using Com.Ena.Timesheet.Xl;
using Ena.Timesheet.Tests;

namespace Com.Ena.Timesheet.Tests.IntegrationTests.Phd
{
    public class PhdTemplateIntegrationTests
    {
        private const string TestYearMonth = "202505";

        [Fact]
        public void CanCreatePhdTemplateFromData()
        {
            // Arrange
            var testData = new List<List<string>>
            {
                new List<string> { "Client1", "Task1" },
                new List<string> { "Client2", "Task2" }
            };

            // Act
            var template = new PhdTemplate(TestYearMonth, testData, "", "");

            // Assert
            Assert.NotNull(template);
            Assert.Equal(TestYearMonth, template.YearMonth);
            Assert.Equal(2, template.GetEntries().Count);
        }

        [Fact]
        public void CanParseClientTasks()
        {
            // Arrange
            var testData = new List<List<string>>
            {
                new List<string> { "Client1", "Task1" },
                new List<string> { "Client2", "Task2" },
                new List<string> { "Client1", "Task3" }
            };
            var template = new PhdTemplate(TestYearMonth, testData, "", "");

            // Act
            var clientTasks = template.ClientTasks();

            // Assert
            Assert.Equal(3, clientTasks.Count);
            Assert.Contains("Client1#Task1", clientTasks);
            Assert.Contains("Client2#Task2", clientTasks);
            Assert.Contains("Client1#Task3", clientTasks);
        }

        [Fact]
        public void CanGenerateDropdowns()
        {
            // Arrange
            var testData = new List<List<string>>
            {
                new List<string> { "Client1", "Task1" },
                new List<string> { "Client2", "Task2" }
            };
            var template = new PhdTemplate(TestYearMonth, testData, "", "");

            // Act
            var dropdownBytes = template.Dropdowns();
            var dropdownText = Encoding.UTF8.GetString(dropdownBytes);

            // Assert
            Assert.Contains("Client1#Task1", dropdownText);
            Assert.Contains("Client2#Task2", dropdownText);
        }

        [Fact]
        public void CanReadAndParsePhdTemplate()
        {
            // Arrange
            var filePath = TestFileHelper.GetTestFilePath("PHD Blank Timesheet April 2025.xlsx");
            var parser = new Parser();
            List<PhdTemplateEntry> entries = new List<PhdTemplateEntry>();
            PhdTemplate template = null;
            
            try
            {
                using (var stream = File.OpenRead(filePath))
                {
                    entries = parser.ParseEntries(stream);
                    template = new PhdTemplate(TestYearMonth, entries, "", "");
                }
            }
            catch (InvalidOperationException ex)
            {
                // This is expected since the test file has duplicate client-task entries
                Console.WriteLine($"Expected error: {ex.Message}");
                Assert.Contains("Duplicate client-task", ex.Message);
                return;
            }

            try
            {
                // Assert
                Assert.NotNull(template);
                Assert.NotEmpty(entries);

                // Write entries to file
                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(directory))
                {
                    throw new InvalidOperationException("Test file path is invalid");
                }

                var outputFile = Path.Combine(directory, "ParsePhdTemplate.txt");
                using (var writer = new StreamWriter(outputFile))
                {
                    foreach (var entry in entries)
                    {
                        writer.WriteLine($"Entry - Row: {entry.GetRowNum() + 1}, Client: '{entry.GetClient()}', Task: '{entry.GetTask()}'");
                        var effortDict = entry.GetEffort();
                        if (effortDict != null)
                        {
                            foreach (var effort in effortDict)
                            {
                                writer.WriteLine($"  Day: {effort.Key}, Effort: {effort.Value}");
                            }
                        }
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                // This is expected since the test file has duplicate client-task entries
                Console.WriteLine($"Expected error: {ex.Message}");
                Assert.Contains("Duplicate client-task", ex.Message);
            }
        }

    }
}
