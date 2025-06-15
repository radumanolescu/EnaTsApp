using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Globalization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;

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
            var sb = new StringBuilder();
            var month = _timesheet.TimesheetMonth;
            
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html lang=\"en\">");
            sb.AppendLine("<head>");
            sb.AppendLine("    <meta charset=\"UTF-8\">");
            sb.AppendLine("    <title>ENA Invoice</title>");
            sb.AppendLine(@"    <style>
        body { font-family: Arial, sans-serif; }
        .invoice-container { max-width: 800px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; margin-bottom: 30px; }
        .invoice-number { font-size: 24px; font-weight: bold; margin: 10px 0; }
        .invoice-date { font-size: 16px; color: #666; }
        .invoice-table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        .invoice-table th, .invoice-table td { 
            border: 1px solid #ddd; 
            padding: 8px; 
            text-align: left;
        }
        .invoice-table th { background-color: #f4f4f4; }
        .total { 
            font-weight: bold; 
            text-align: right;
            padding: 10px;
            border-top: 2px solid #ddd;
        }
    </style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("    <div class=\"invoice-container\">");
            sb.AppendLine("        <div class=\"header\">");
            sb.AppendLine($"            <h1>ENA Invoice</h1>");
            sb.AppendLine($"            <div class=\"invoice-number\">Invoice for {month.ToString("MMMM yyyy")}</div>");
            sb.AppendLine($"            <div class=\"invoice-date\">Generated on: {DateTime.Now.ToString("yyyy-MM-dd")}</div>");
            sb.AppendLine("        </div>");

            sb.AppendLine("        <table class=\"invoice-table\">");
            sb.AppendLine("            <thead>");
            sb.AppendLine("                <tr>");
            sb.AppendLine("                    <th>Date</th>");
            sb.AppendLine("                    <th>Project</th>");
            sb.AppendLine("                    <th>Hours</th>");
            sb.AppendLine(@"                    <th>Description</th>
                </tr>");
            sb.AppendLine("            </thead>");
            sb.AppendLine("            <tbody>");

            foreach (var entry in _timesheet.EnaTsEntries.OrderBy(e => e.Date))
            {
                sb.AppendLine("                <tr>");
                sb.AppendLine($"                    <td>{entry.Date.ToString("yyyy-MM-dd")}</td>");
                sb.AppendLine($"                    <td>{entry.ProjectId}</td>");
                sb.AppendLine($"                    <td>{entry.Hours}</td>");
                sb.AppendLine($"                    <td>{entry.Description}</td>");
                sb.AppendLine("                </tr>");
            }

            sb.AppendLine("            </tbody>");
            sb.AppendLine("        </table>");

            var totalHours = _timesheet.EnaTsEntries.Sum(e => e.Hours);
            sb.AppendLine($"        <div class=\"total\">Total Hours: {totalHours}</div>");
            sb.AppendLine("    </div>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
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
