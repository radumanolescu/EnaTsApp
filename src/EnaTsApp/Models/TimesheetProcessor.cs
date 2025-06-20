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

        public static string GetDownloadsDirectory()
        {
            //return Environment.GetFolderPath(Environment.KnownFolders.Downloads);
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            //return KnownFolders.GetPath(KnownFolder.Downloads);
        }

        public string Process()
        {
            try
            {
                _logger.LogInformation($"Starting timesheet processing for {_yyyyMM}");
                
                List<List<string>>? templateData;
                List<List<string>>? timesheetData;

                _logger.LogInformation("Parsing Excel files");
                templateData = ParseExcelFile(_templatePath);
                timesheetData = ParseExcelFile(_timesheetPath);

                if (templateData == null || templateData.Count == 0 )
                {
                    _logger.LogError("Failed to parse template file");
                    throw new Exception("Failed to parse template");
                }
                if (timesheetData == null || timesheetData.Count == 0)
                {
                    _logger.LogError("Failed to parse timesheet file");
                    throw new Exception("Failed to parse timesheet");
                }

                _logger.LogInformation("Creating PHD template and ENA timesheet objects");
                // Parse the PHD template and the ENA timesheet
                var phdTemplate = new PhdTemplate(_yyyyMM, templateData, _templatePath, GetTemplateOutputPath(_yyyyMM));
                var enaTimesheet = new EnaTimesheet(_yyyyMM, timesheetData, _timesheetPath, GetTimesheetOutputPath());
                // Generate the HTML invoice based on the ENA timesheet
                EnaInvoice enaInvoice = new EnaInvoice(enaTimesheet);
                var invoiceHtml = enaInvoice.GenerateInvoiceHtml();
                File.WriteAllText("invoice.html", invoiceHtml);
                // Generate the drop-downs for the following month based on the template
                WriteDropdowns(phdTemplate);                

                _logger.LogInformation("Updating PHD template with ENA timesheet data");
                phdTemplate.Update(enaTimesheet);
                phdTemplate.SaveAs();

                _logger.LogInformation("Timesheet processing completed successfully");
                return phdTemplate.OutputPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during timesheet processing");
                throw;
            }
        }

        public string GetTemplateOutputPath(string yyyyMM)
        {
            return Path.Combine(GetDownloadsDirectory(), PhdTemplate.GetTimesheetFileName(yyyyMM));
        }

        public string GetTimesheetOutputPath()
        {
            return Path.Combine(GetDownloadsDirectory(), EnaTimesheet.AddRevisionToFilename(Path.GetFileName(_timesheetPath)));
        }

        public void WriteDropdowns(PhdTemplate phdTemplate)
        {
            try
            {
                _logger.LogInformation("Writing client tasks to Dropdowns.txt");
                var dropdownsPath = Path.Combine(Path.GetDirectoryName(_templatePath) ?? throw new InvalidOperationException("Invalid template path"), "Dropdowns.txt");
                var tasks = phdTemplate.ClientTasks();
                File.WriteAllLines(dropdownsPath, tasks);
                _logger.LogInformation($"Client tasks written to {dropdownsPath}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error writing client tasks to file");
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

    }
}

