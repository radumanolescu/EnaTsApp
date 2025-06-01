using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Com.Ena.Timesheet.Xl;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Ena;
using OfficeOpenXml;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;

namespace Com.Ena.Timesheet
{
    public class TimesheetProcessor
    {
        private readonly string _templatePath;
        private readonly string _timesheetPath;
        private readonly string _yyyyMM;
        private readonly ILogger<TimesheetProcessor> _logger;

        public TimesheetProcessor(string yyyyMM, string templatePath, string timesheetPath)
        {
            _yyyyMM = yyyyMM;
            _templatePath = templatePath;
            _timesheetPath = timesheetPath;
            _logger = CreateLogger();
            _logger.LogInformation($"Created TimesheetProcessor for {_yyyyMM}");
            _logger.LogInformation($"Template path: {_templatePath}");
            _logger.LogInformation($"Timesheet path: {_timesheetPath}");
        }

        private ILogger<TimesheetProcessor> CreateLogger()
        {
            var serviceProvider = new ServiceCollection()
                .AddLogging(builder => builder.AddConsole())
                .BuildServiceProvider();
            return serviceProvider.GetRequiredService<ILogger<TimesheetProcessor>>();
        }

        public void Process()
        {
            try
            {
                _logger.LogInformation($"Starting timesheet processing for {_yyyyMM}");
                
                List<List<string>>? templateData;
                List<List<string>>? timesheetData;

                _logger.LogInformation("Parsing Excel files");
                templateData = ParseExcelFile(_templatePath);
                timesheetData = ParseExcelFile(_timesheetPath);

                if (templateData == null)
                {
                    _logger.LogError("Failed to parse template file");
                    throw new Exception("Failed to parse template");
                }
                if (timesheetData == null)
                {
                    _logger.LogError("Failed to parse timesheet file");
                    throw new Exception("Failed to parse timesheet");
                }

                _logger.LogInformation("Creating PHD template and ENA timesheet objects");
                var phdTemplate = new PhdTemplate(_yyyyMM, templateData, _templatePath, PhdTimesheetFileName(_yyyyMM));
                var enaTimesheet = new EnaTimesheet(_yyyyMM, timesheetData, _timesheetPath, AddRevisionToFilename(_timesheetPath));

                _logger.LogInformation("Updating PHD template with ENA timesheet data");
                phdTemplate.Update(enaTimesheet);

                _logger.LogInformation("Timesheet processing completed successfully");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during timesheet processing");
                throw;
            }
        }

        private List<List<string>> ParseExcelFile(string filePath)
        {
            try
            {
                var parser = new ExcelParser();
                var data = parser.ParseExcelFile(filePath);
                if (data != null)
                {
                    Console.WriteLine($"Successfully parsed Excel file: {data.Count} rows, {data[0].Count} columns");
                    return data;
                }
                return new List<List<string>>();
            }
            catch (InvalidOperationException ex)
            {
                Console.WriteLine(ex.Message);
                return new List<List<string>>();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error parsing Excel file: {ex.Message}");
                return new List<List<string>>();
            }
        }

        private string PhdTimesheetFileName(string yyyyMM)
        {
            // e.g. "PHD ENA Timesheet 2023-03.xlsx"
            return "PHD ENA Timesheet " + yyyyMM.Substring(0, 4) + "-" + yyyyMM.Substring(4, 2) + ".xlsx";
        }

        /// <summary>
        /// Adds 'R' before the extension in an Excel filename.
        /// For example, "PHD 04 - April 2025.xlsx" becomes "PHD 04 - April 2025R.xlsx"
        /// </summary>
        public static string AddRevisionToFilename(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentException("Filename cannot be null or empty", nameof(filename));

            string extension = Path.GetExtension(filename);
            if (string.IsNullOrEmpty(extension))
                throw new ArgumentException("Filename must have an extension", nameof(filename));

            string baseName = Path.GetFileNameWithoutExtension(filename);
            return $"{baseName}R{extension}";
        }

    }
}

