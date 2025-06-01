using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Com.Ena.Timesheet.Phd;
using EnaTsApp.Tests.TestHelpers;
using Com.Ena.Timesheet.Xl;
using OfficeOpenXml;
using Ena.Timesheet.Tests;

namespace Com.Ena.Timesheet.Tests.IntegrationTests.Phd
{
    public class PhdTemplateIntegrationTests
    {
        public PhdTemplateIntegrationTests()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        }

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
            var parser = new ExcelParser();
            List<List<string>> entries = parser.ParseExcelFile(filePath);
            PhdTemplate template = new PhdTemplate(TestYearMonth, entries, filePath, filePath);

            // Assert
            Assert.NotNull(template);
            Assert.NotEmpty(template.GetEntries());

            // Write entries to file
            var directory = Path.GetDirectoryName(filePath);
            if (string.IsNullOrEmpty(directory))
            {
                throw new InvalidOperationException("Test file path is invalid");
            }

            var outputFile = Path.Combine(directory, "ParsePhdTemplate.txt");
            using (var writer = new StreamWriter(outputFile))
            {
                foreach (var entry in template.GetEntries())
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

    }
}
