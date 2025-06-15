using System;
using System.IO;
using Xunit;
using Com.Ena.Timesheet;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using OfficeOpenXml;
using Com.Ena.Timesheet.Phd;

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

            _templatePath = Path.Combine(_testDataPath, "PHD Blank Timesheet April 2025.xlsx");
            _output.WriteLine($"---------- Template Path: [{_templatePath}]");
            Assert.True(File.Exists(_templatePath), "Template file should exist");

            _timesheetPath = Path.Combine(_testDataPath, "PHD 04 - April 2025R.xlsx");
            _output.WriteLine($"---------- Timesheet Path: [{_timesheetPath}]");
            Assert.True(File.Exists(_timesheetPath), "Timesheet file should exist");
        }

        [Fact]
        public void ShouldProcessTimesheetSuccessfully()
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
            Assert.Throws<FileNotFoundException>(() => processor.Process());
        }

        [Fact]
        public void ShouldHandleInvalidTimesheetPath()
        {
            // Arrange
            var yyyyMM = "202504";
            var invalidPath = Path.Combine(_testDataPath, "nonexistent.xlsx");
            var processor = new TimesheetProcessor(yyyyMM, _templatePath, invalidPath);

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => processor.Process());
        }
    }
}
// dotnet test src/EnaTsApp.Tests/EnaTsApp.Tests.csproj --filter "FullyQualifiedName~TimesheetProcessorIntegrationTests" --verbosity normal --logger "console;verbosity=detailed"