using System;
using System.IO;
using OfficeOpenXml;

namespace Com.Ena.Timesheet
{
    public class ExcelMapped
    {
        private readonly string _inputPath;
        private readonly string _outputPath;
        protected readonly ExcelPackage _excelPackage;

        public ExcelMapped(string inputPath, string outputPath)
        {
            _inputPath = inputPath;
            _outputPath = outputPath;

            // If either path is empty, return early without initialization
            if (string.IsNullOrEmpty(inputPath) || string.IsNullOrEmpty(outputPath))
            {
                return;
            }

            // Validate paths
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("Input file not found", inputPath);
            }

            // Create output directory if it doesn't exist
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Initialize ExcelPackage
            _excelPackage = new ExcelPackage(new FileInfo(inputPath));
        }

        /// <summary>
        /// Saves the ExcelPackage to the output path specified in the constructor.
        /// </summary>
        public void SaveAs()
        {
            if (_excelPackage != null)
            {
                _excelPackage.SaveAs(new FileInfo(_outputPath));
            }
        }
    }
}
