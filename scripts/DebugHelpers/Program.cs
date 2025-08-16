using System;
using System.IO;
using System.Linq;
using Com.Ena.Timesheet;
using Microsoft.Extensions.Logging;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("TimesheetProcessor Debug Helper");
        Console.WriteLine("=============================\n");

        try
        {
            // Get input paths from command line or use defaults
            string yyyyMM = args.Length > 0 ? args[0] : DateTime.Now.ToString("yyyyMM");
            string templatePath = args.Length > 1 ? args[1] : @"C:\path\to\your\template.xlsx";
            string timesheetPath = args.Length > 2 ? args[2] : @"C:\path\to\your\timesheet.xlsx";

            Console.WriteLine($"Using parameters:");
            Console.WriteLine($"- Date (YYYYMM): {yyyyMM}");
            Console.WriteLine($"- Template: {templatePath}");
            Console.WriteLine($"- Timesheet: {timesheetPath}\n");

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
                }

                // Show available client tasks
                Console.WriteLine("\nAvailable client tasks:");
                foreach (var task in processor.ClientTasks.Take(20)) // Show first 20 to avoid flooding
                {
                    Console.WriteLine($"- {task}");
                }
                if (processor.ClientTasks.Count > 20)
                {
                    Console.WriteLine($"... and {processor.ClientTasks.Count - 20} more");
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
