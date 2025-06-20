using System;
using System.IO;
using System.Linq;
using System.Globalization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using Cottle;

namespace Com.Ena.Timesheet.Ena
{
    public class EnaInvoice
    {
        private readonly ILogger<EnaInvoice> _logger;
        private readonly EnaTimesheet _timesheet;

        public EnaInvoice(EnaTimesheet timesheet)
        {
            _logger = GetLogger<EnaInvoice>();
            _timesheet = timesheet;
        }

        public string GenerateInvoiceHtml()
        {
            DateTime month = _timesheet.TimesheetMonth;
            string formattedMonth = month.ToString("MMMM yyyy", CultureInfo.InvariantCulture);
            DateTime date = DateTime.Now;
            string formattedDate = date.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture);  // "June 12, 2025"
            List<EnaTsEntry> entries = _timesheet.EnaTsEntries.OrderBy(e => e.Date).ThenBy(e => e.ProjectId).ToList();
            string entriesWithTotals = _timesheet.GetEntriesWithTotalsAsHtml();
            string projectEntries = _timesheet.GetProjectEntriesAsHtml();
            float? totalHours = entries.Sum(e => e.Hours);

            string template = File.ReadAllText(GetInvoiceTemplatePath());

            DocumentResult documentResult = Document.CreateDefault(template); // Create from template string
            var document = documentResult.DocumentOrThrow; // Throws ParseException on error
            var context = Context.CreateBuiltin(new Dictionary<Value, Value>
            {
                ["invoiceMonth"] = formattedMonth,
                ["invoiceDate"] = formattedDate,
                ["entriesWithTotals"] = entriesWithTotals,
                ["projectEntries"] = projectEntries
            });
            string html = document.Render(context);
            File.WriteAllText(GetInvoicePath(month), html);
            return html;
        }

        private static string GetInvoiceTemplatePath()
        {
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "Templates", "ena-invoice.html");
            if (!File.Exists(templatePath))
            {
                // If not found, try relative to the project root
                templatePath = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).FullName, "Resources", "Templates", "ena-invoice.html");
            }
            // Until we know how to reliably access this resource, return the absolute path
            return "C:/Users/Radu/-/projects/C#/EnaTsUiV2/src/EnaTsApp/Resources/Templates/ena-invoice.html";
        }

        private static string GetInvoicePath(DateTime month)
        {
            string formattedMonth =$"ENA Invoice {month:yyyy-MM}.html";  // e.g.: "ENA Invoice 2025-04.html"
            return Path.Combine(TimesheetProcessor.GetDownloadsDirectory(), formattedMonth);
        }

        private static ILogger<T> GetLogger<T>()
        {
            var serviceProvider = new ServiceCollection()
                .AddLogging(builder => builder.AddConsole())
                .BuildServiceProvider();
            return serviceProvider.GetRequiredService<ILogger<T>>();
        }
    }
}
