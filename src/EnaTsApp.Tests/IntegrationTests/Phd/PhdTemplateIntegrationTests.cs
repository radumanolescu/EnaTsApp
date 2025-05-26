using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;
using Com.Ena.Timesheet.Phd;

namespace EnaTsApp.Tests.IntegrationTests.Phd
{
    public class PhdTemplateIntegrationTests
    {
        private const string TestYearMonth = "202305";

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
            var template = new PhdTemplate(TestYearMonth, testData);

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
            var template = new PhdTemplate(TestYearMonth, testData);

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
            var template = new PhdTemplate(TestYearMonth, testData);

            // Act
            var dropdownBytes = template.Dropdowns();
            var dropdownText = Encoding.UTF8.GetString(dropdownBytes);

            // Assert
            Assert.Contains("Client1#Task1", dropdownText);
            Assert.Contains("Client2#Task2", dropdownText);
        }
    }
}
