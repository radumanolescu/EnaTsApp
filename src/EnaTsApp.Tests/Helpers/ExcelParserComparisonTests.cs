using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Com.Ena.Timesheet.Xl;
using Xunit;
using Xunit.Abstractions;
using OfficeOpenXml;

namespace Com.Ena.Timesheet.Tests.Helpers
{
    public class ExcelParserComparisonTests
    {
        private static string _testDataPath = Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "TestData");
        private readonly ITestOutputHelper _output;

        public ExcelParserComparisonTests(ITestOutputHelper output)
        {
            _output = output;
            // Set up EPPlus license context
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        }

        public static IEnumerable<object[]> GetTestExcelFiles()
        {
            // Add paths to test Excel files here
            yield return new object[] { Path.Combine(_testDataPath, "ENA-TimesheetFragment.xlsx") };
            // Add more test files as needed
        }

        [Theory]
        [MemberData(nameof(GetTestExcelFiles))]
        public void ExcelParsers_ShouldProduceSameResults(string excelFilePath)
        {
            if (!File.Exists(excelFilePath))
            {
                throw new FileNotFoundException($"Test file not found: {excelFilePath}");
            }

            // Arrange
            var npoiParser = new ExcelParserNPOI();
            var epplusParser = new ExcelParser();

            // Act
            var npoiResult = npoiParser.ParseExcelFile(excelFilePath);
            var epplusResult = epplusParser.ParseExcelFile(excelFilePath);

            // Assert - compare row counts
            Assert.Equal(npoiResult.Count, epplusResult.Count);

            // Compare each cell
            for (int row = 0; row < npoiResult.Count; row++)
            {
                var npoiRow = npoiResult[row];
                var epplusRow = epplusResult[row];

                // Compare column counts for the row
                Assert.Equal(npoiRow.Count, epplusRow.Count);

                for (int col = 0; col < npoiRow.Count; col++)
                {
                    var npoiValue = npoiRow[col] ?? string.Empty;
                    var epplusValue = epplusRow[col] ?? string.Empty;

                    // Normalize values for comparison
                    npoiValue = NormalizeValue(npoiValue);
                    epplusValue = NormalizeValue(epplusValue);

                    // Log the first few differences for debugging
                    if (npoiValue != epplusValue && row < 5 && col < 5)
                    {
                        _output.WriteLine($"Difference at [{row},{col}]:");
                        _output.WriteLine($"NPOI: '{npoiRow[col]}'");
                        _output.WriteLine($"EPPlus: '{epplusRow[col]}'");
                    }

                    Assert.Equal(npoiValue, epplusValue);
                }
            }
        }

        private string NormalizeValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            // Trim whitespace and normalize line endings
            return value.Trim()
                       .Replace("\r\n", "\n")
                       .Replace("\r", "\n");
        }
    }
}
