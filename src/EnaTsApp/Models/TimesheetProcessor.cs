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
        private HashSet<string> clientTaskSet;
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

        /// <summary>
        /// Processes the timesheet data by:
        /// 1. Parsing both the template and timesheet Excel files
        /// 2. Creating PHD template and ENA timesheet objects
        /// 3. Updating the PHD template with ENA timesheet data
        /// 4. Generating an HTML invoice based on the ENA timesheet
        /// </summary>
        /// <returns>The output path of the processed ENA timesheet file.</returns>
        /// <exception cref="System.Exception">Thrown if either file parsing fails or the PHD template update fails.</exception>
        public string Process()
        {
            try
            {
                _logger.LogInformation($"Starting timesheet processing for {_yyyyMM}");
                
                List<List<string>> templateData;
                List<List<string>> timesheetData;

                _logger.LogInformation("Parsing Excel files");
                templateData = ParseExcelFile(_templatePath);
                timesheetData = ParseExcelFile(_timesheetPath);

                if (templateData.Count == 0 )
                {
                    _logger.LogError("Failed to parse template file");
                    throw new Exception("Failed to parse template");
                }
                if (timesheetData.Count == 0)
                {
                    _logger.LogError("Failed to parse timesheet file");
                    throw new Exception("Failed to parse timesheet");
                }

                _logger.LogInformation("Creating PHD template and ENA timesheet objects");
                // Parse the PHD template and the ENA timesheet
                var phdTemplate = new PhdTemplate(_yyyyMM, templateData, _templatePath, GetTemplateOutputPath(_yyyyMM));
                var enaTimesheet = new EnaTimesheet(_yyyyMM, timesheetData, _timesheetPath, GetTimesheetOutputPath());
                
                // Generate the drop-downs for the following month based on the template
                WriteDropdowns(phdTemplate);

                // Fill in information from the ENA timesheet into the PHD template
                _logger.LogInformation("Updating PHD template with ENA timesheet data");
                bool valid = phdTemplate.Update(enaTimesheet);
                phdTemplate.SaveAs();
                if (!valid)
                {
                    var errorMessage = "See error annotations in the timesheet file " + enaTimesheet.OutputPath;
                    _logger.LogError(errorMessage);
                    throw new Exception(errorMessage);
                }

                // Generate the HTML invoice based on the ENA timesheet
                EnaInvoice enaInvoice = new EnaInvoice(enaTimesheet);
                var invoiceHtml = enaInvoice.GenerateInvoiceHtml();
                File.WriteAllText("invoice.html", invoiceHtml);

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
                _logger.LogInformation("Writing client tasks to ena_dropdown.txt");
                var dropdownsPath = Path.Combine(Path.GetDirectoryName(phdTemplate.OutputPath) ?? throw new InvalidOperationException("Invalid template path"), "ena_dropdown.txt");
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


        /// <summary>
        /// Validates timesheet entries against a template and returns a dictionary of invalid activities
        /// along with their suggested best matches from the template.
        /// </summary>
        /// <returns>A dictionary where keys are invalid activities and values are the suggested best matches</returns>
        public Dictionary<string, string> Validate()
        {
            _logger.LogInformation("Validating timesheet entries against template");
            var invalidActivities = new Dictionary<string, string>();
            
            // Parse the timesheet data
            var timesheetData = ParseExcelFile(_timesheetPath);
            if (timesheetData == null || timesheetData.Count == 0)
            {
                _logger.LogError("Failed to parse timesheet file for validation");
                throw new Exception("Failed to parse timesheet for validation");
            }

            // Create a timesheet instance without saving it
            var enaTimesheet = new EnaTimesheet(_yyyyMM, timesheetData, _timesheetPath, string.Empty);
            List<List<string>> templateData = ParseExcelFile(_templatePath);
            var phdTemplate = new PhdTemplate(_yyyyMM, templateData, _templatePath, GetTemplateOutputPath(_yyyyMM));
            clientTaskSet = new HashSet<string>(phdTemplate.ClientTasks());

            // Check each entry
            foreach (var entry in enaTimesheet.GetEntries())
            {
                string projectActivity = entry.ProjectActivity();
                if (!clientTaskSet.Contains(projectActivity))
                {
                    string suggested = EnaTimesheet.BestMatch(projectActivity, clientTaskSet);
                    if (!string.IsNullOrEmpty(suggested) && !invalidActivities.ContainsKey(projectActivity))
                    {
                        invalidActivities[projectActivity] = suggested;
                        _logger.LogWarning($"Found invalid activity: '{projectActivity}'. Suggested: '{suggested}'");
                    }
                }
            }

            _logger.LogInformation($"Validation complete. Found {invalidActivities.Count} invalid activities");
            return invalidActivities;
        }
    }
}

