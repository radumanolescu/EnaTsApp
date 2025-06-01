using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Com.Ena.Timesheet.Xl;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Ena;
using OfficeOpenXml;
using System.IO;

namespace Com.Ena.Timesheet
{
    public class TimesheetProcessor
    {
        private readonly string _templatePath;
        private readonly string _timesheetPath;
        private readonly string _yyyyMM;

        public TimesheetProcessor(string yyyyMM, string templatePath, string timesheetPath)
        {
            _yyyyMM = yyyyMM;
            _templatePath = templatePath;
            _timesheetPath = timesheetPath;
        }

        public void Process()
        {
            List<List<string>>? templateData;
            List<List<string>>? timesheetData;

            templateData = ParseExcelFile(_templatePath);
            timesheetData = ParseExcelFile(_timesheetPath);

            if (templateData == null)
            {
                throw new Exception("Failed to parse template");
            }
            if (timesheetData == null)
            {
                throw new Exception("Failed to parse timesheet");
            }

            var phdTemplate = new PhdTemplate(_yyyyMM, templateData);
            var enaTimesheet = new EnaTimesheet(_yyyyMM, timesheetData);

            phdTemplate.Update(enaTimesheet);

            // ToDo: implement
            //var xlsxBytes = phdTemplate.UpdateXlsx(0);
            //string phdTimesheetFileName = PhdTimesheetFileName(_yyyyMM);

            //File.WriteAllBytes(phdTimesheetFileName, xlsxBytes);
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
    }
}

