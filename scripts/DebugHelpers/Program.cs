using System;
using System.IO;
using System.Linq;
using Com.Ena.Timesheet;
using Com.Ena.Timesheet.Ena;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        TestActivityUpdateFlow(args);
    }
    
    /// <summary>
    /// Tests the activity update flow by validating a timesheet and displaying invalid activities.
    /// Allows updating invalid activities in the Excel file with valid ones.
    /// </summary>
    /// <param name="args">Command line arguments in the format: [YYYYMM] [templatePath] [timesheetPath]</param>
    /// <remarks>
    /// If no arguments are provided, default paths will be used:
    /// - Current date for YYYYMM
    /// - Default template and timesheet paths
    /// </remarks>
    static void TestActivityUpdateFlow(string[] args)
    {
        Console.WriteLine("TimesheetProcessor Debug Helper");
        Console.WriteLine("=============================\n");

        try
        {
            // Get input paths from command line or use defaults
            string yyyyMM = args.Length > 0 ? args[0] : DateTime.Now.ToString("yyyyMM");
            string templatePath = args.Length > 1 ? args[1] : @"C:\Users\Radu\-\projects\C#\EnaTsUiV2\src\EnaTsApp.Tests\TestData\PHD Blank Timesheet April 2025.xlsx";
            string timesheetPath = args.Length > 2 ? args[2] : @"C:\Users\Radu\-\projects\C#\EnaTsUiV2\src\EnaTsApp.Tests\TestData\ENA-TimesheetFragment.xlsx";

            Console.WriteLine($"Using parameters:");
            Console.WriteLine($"- Date (YYYYMM): {yyyyMM}");
            Console.WriteLine($"- Template: {templatePath}");
            Console.WriteLine($"- Timesheet: {timesheetPath}\n");

            List<List<string>> timesheetData = TimesheetProcessor.ParseExcelFile(timesheetPath);
            var enaTimesheet = new EnaTimesheet(yyyyMM, timesheetData, timesheetPath, string.Empty);

            // Create and validate the processor
            Console.WriteLine("Creating TimesheetProcessor...");
            var processor = new TimesheetProcessor(yyyyMM, templatePath, timesheetPath);
            
            // Validate the timesheet
            Console.WriteLine("\nValidating timesheet...");
            var invalidActivities = processor.Validate();
            
            if (invalidActivities.Count > 0)
            {
                Console.WriteLine($"\nFound {invalidActivities.Count} invalid activities:");
                foreach (var (invalid, suggested) in invalidActivities)
                {
                    Console.WriteLine($"- Invalid: {invalid}");
                    Console.WriteLine($"  Suggested: {suggested}");
                    // Update the activity in the timesheet
                    // Note: The UpdateActivity method expects the full project#activity string
                    // If 'invalid' is just the activity part, we need to get the full string from the entry
                    enaTimesheet.UpdateActivity(invalid, suggested);
                }

                // Show available client tasks
                Console.WriteLine("\nAvailable client tasks:");
                int maxShowTasks = 10;
                foreach (var task in processor.ClientTasks.Take(maxShowTasks)) // Show first N to avoid flooding
                {
                    Console.WriteLine($"- {task}");
                }
                if (processor.ClientTasks.Count > maxShowTasks)
                {
                    Console.WriteLine($"... and {processor.ClientTasks.Count - maxShowTasks} more");
                }
            }
            else
            {
                Console.WriteLine("\nNo validation errors found in the timesheet.");
                
                // Try processing if no errors
                Console.WriteLine("\nAttempting to process timesheet...");
                string outputPath = processor.Process();
                Console.WriteLine($"\nTimesheet processed successfully!");
                Console.WriteLine($"Output file: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("\nERROR:");
            Console.WriteLine(ex.ToString());
        }

        // Console.WriteLine("\nPress any key to exit...");
        // Console.ReadKey();
    }
}
