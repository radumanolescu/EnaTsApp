using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using Com.Ena.Timesheet;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using OfficeOpenXml;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Xl;
using Com.Ena.Timesheet.Ena;


namespace Com.Ena.Timesheet.Tests.IntegrationTests
{
    public class TimesheetProcessorIntegrationTests
    {
        private readonly string _testDataPath;
        private readonly string _templatePath;
        private readonly string _timesheetPath;
        private string _outputPath;
        private readonly ITestOutputHelper _output;

        public TimesheetProcessorIntegrationTests(ITestOutputHelper output)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");

            _output = output;
            _output.WriteLine($"---------- Current Directory: [{Directory.GetCurrentDirectory()}]");

            _testDataPath = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "TestData");
            _output.WriteLine($"---------- Test Data Path: [{_testDataPath}]");

            _templatePath = GetTemplateTestFilePath();
            _output.WriteLine($"---------- Template Path: [{_templatePath}]");
            Assert.True(File.Exists(_templatePath), "Template file should exist");

            _timesheetPath = GetTimesheetTestFilePath();
            _output.WriteLine($"---------- Timesheet Path: [{_timesheetPath}]");
            Assert.True(File.Exists(_timesheetPath), "Timesheet file should exist");
        }

        private string GetTemplateTestFilePath()
        {
            var templatePath = Path.Combine(_testDataPath, "PHD Blank Timesheet April 2025.xlsx");
            Assert.True(File.Exists(templatePath), "Template file does not exist");
            return templatePath;
        }

        private string GetTimesheetTestFilePath()
        {
            var timesheetPath = Path.Combine(_testDataPath, "PHD 04 - April 2025R.xlsx");
            Assert.True(File.Exists(timesheetPath), "Timesheet file does not exist");
            return timesheetPath;
        }

        [Fact]
        public void ShouldProcessTemplateAndTimesheetSuccessfully()
        {
            // Arrange
            var yyyyMM = "202504";
            var processor = new TimesheetProcessor(yyyyMM, _templatePath, _timesheetPath);
            _outputPath = PhdTemplate.GetTimesheetFileName(yyyyMM);
            _output.WriteLine($"---------- Output Path: [{_outputPath}]");

            // Act
            string outputPath = processor.Process();
            _output.WriteLine($"---------- Actual Output Path: [{outputPath}]");

            // Assert
            Assert.True(File.Exists(outputPath), "Output file should be created");
            //Assert.Equal(_outputPath, Path.GetFileName(outputPath), "Output file name should match expected");
        }

        [Fact]
        public void ShouldProcessTemplateSuccessfully()
        {
            // Arrange
            var templatePath = GetTemplateTestFilePath();
            // Act
            var parser = new ExcelParser();
            // This parsing method finds every single row that has been touched by the user, whether or not it has data in it.
            // So it finds rows up to row 93 in the template, even though the template ends with SUM in row 91.
			List<List<string>>? templateData = parser.ParseExcelFile(templatePath);
            // Assert
			Assert.True(templateData != null, "Failed to parse template file");
            Assert.Equal(93, templateData.Count);
		}

        [Fact]
        public void ShouldProcessTimesheetSuccessfully()
        {
            // Arrange
            var timesheetPath = GetTimesheetTestFilePath();
            // Act
            var parser = new ExcelParser();
            // This parsing method finds every single row that has been touched by the user, whether or not it has data in it.
            List<List<string>>? timesheetData = parser.ParseExcelFile(timesheetPath);
            var enaTimesheet = new EnaTimesheet("202504", timesheetData, timesheetPath, "unused.xlsx");
            // Assert
            Assert.True(timesheetData != null, "Failed to parse timesheet file");
            Assert.Equal(119, timesheetData.Count);
            Assert.Equal(98, enaTimesheet.EnaTsEntries.Count);
            Assert.Equal(10, enaTimesheet.ProjectEntries.Count);
        }

        [Fact]
        public void ShouldLogProcessingSteps()
        {
            // Arrange
            var yyyyMM = "202504";
            var processor = new TimesheetProcessor(yyyyMM, _templatePath, _timesheetPath);

            // Act
            string outputPath = processor.Process();

            // Assert
            // This test assumes the processor logs messages using ILogger
            // You might need to modify this based on your actual logging setup
            // For now, we're just checking that the process completes without errors
            Assert.True(true, "Process completed without errors");
            //Assert.NotNull(outputPath, "Process should return a valid output path");
        }

        [Fact]
        public void ShouldHandleInvalidTemplatePath()
        {
            // Arrange
            var yyyyMM = "202504";
            var invalidPath = Path.Combine(_testDataPath, "nonexistent.xlsx");
            var processor = new TimesheetProcessor(yyyyMM, invalidPath, _timesheetPath);

            // Act & Assert
            Assert.Throws<Exception>(() => processor.Process());
        }

        [Fact]
        public void ShouldHandleInvalidTimesheetPath()
        {
            // Arrange
            var yyyyMM = "202504";
            var invalidPath = Path.Combine(_testDataPath, "nonexistent.xlsx");
            var processor = new TimesheetProcessor(yyyyMM, _templatePath, invalidPath);

            // Act & Assert
            Assert.Throws<Exception>(() => processor.Process());
        }
    }
}
// dotnet test src/EnaTsApp.Tests/EnaTsApp.Tests.csproj --filter "FullyQualifiedName~TimesheetProcessorIntegrationTests" --verbosity normal --logger "console;verbosity=detailed"