using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using Com.Ena.Timesheet.Phd;

namespace EnaTsApp
{
    public class MainForm : Form
    {
        private Button? btnChooseMonth;
        private Button? btnUploadTemplate;
        private Button? btnUploadTimesheet;
        private Button? btnProcess;
        private Panel? errorPanel;
        private Label? errorLabel;
        private DataGridView? dataGrid;
        
        private DateTime selectedDate = DateTime.Now;
        private List<List<string>>? templateData;
        private List<List<string>>? timesheetData;
        private PhdTemplate? phdTemplate;

        public MainForm()
        {
            InitializeComponent();
            // Set EPPlus license context for non-commercial use
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage.License.SetNonCommercialPersonal("Elaine Newman");
        }

        private void InitializeComponent()
        {
            // Initialize all controls to null
            btnChooseMonth = null;
            btnUploadTemplate = null;
            btnUploadTimesheet = null;
            btnProcess = null;
            errorPanel = null;
            errorLabel = null;
            dataGrid = null;
            templateData = null;
            timesheetData = null;

            this.SuspendLayout();

            // Form properties
            this.Text = "Elaine Newman Architect, PC - Timesheet & Invoice App";
            this.Size = new Size(1800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Create buttons panel
            Panel buttonPanel = new Panel
            {
                Height = 50,
                Dock = DockStyle.Top,
                Padding = new Padding(10, 10, 10, 10)
            };

            // Button 1: Choose Month
            btnChooseMonth = new Button
            {
                Text = "Choose Month",
                Size = new Size(120, 30),
                Location = new Point(10, 10)
            };
            btnChooseMonth.Click += BtnChooseMonth_Click;

            // Button 2: Upload Template
            btnUploadTemplate = new Button
            {
                Text = "Upload Template",
                Size = new Size(120, 30),
                Location = new Point(140, 10)
            };
            btnUploadTemplate.Click += BtnUploadTemplate_Click;

            // Button 3: Upload Timesheet
            btnUploadTimesheet = new Button
            {
                Text = "Upload Timesheet",
                Size = new Size(120, 30),
                Location = new Point(270, 10)
            };
            btnUploadTimesheet.Click += BtnUploadTimesheet_Click;

            // Button 4: Process (optional fourth button)
            btnProcess = new Button
            {
                Text = "Process Data",
                Size = new Size(120, 30),
                Location = new Point(400, 10)
            };
            btnProcess.Click += BtnProcess_Click;

            buttonPanel.Controls.AddRange(new Control[] { 
                btnChooseMonth, btnUploadTemplate, btnUploadTimesheet, btnProcess 
            });

            // Error panel
            errorPanel = new Panel
            {
                Height = 80,
                Dock = DockStyle.Top,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(255, 245, 245)
            };

            Label errorTitle = new Label
            {
                Text = "Errors / Messages",
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                Location = new Point(10, 5),
                Size = new Size(300, 20)
            };

            errorLabel = new Label
            {
                Text = "No errors",
                Location = new Point(10, 25),
                Size = new Size(850, 50),
                ForeColor = Color.Green,
                AutoSize = false
            };

            errorPanel.Controls.AddRange(new Control[] { errorTitle, errorLabel });

            // Data grid
            dataGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            // Add controls to form
            this.Controls.Add(dataGrid);
            this.Controls.Add(errorPanel);
            this.Controls.Add(buttonPanel);

            this.ResumeLayout(false);
        }

        private void BtnChooseMonth_Click(object? sender, EventArgs e)
        {
            try
            {
                using MonthCalendar calendar = new MonthCalendar();
                Form calendarForm = new Form
                {
                    Text = "Choose Date",
                    Size = new Size(300, 250),
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false
                };

                calendar.Location = new Point(10, 10);
                calendar.DateSelected += (s, args) =>
                {
                    if (calendarForm != null)
                    {
                        selectedDate = calendar.SelectionStart;
                        calendarForm.DialogResult = DialogResult.OK;
                        calendarForm.Close();
                    }
                };

                Button okButton = new Button
                {
                    Text = "OK",
                    Size = new Size(75, 23),
                    Location = new Point(100, 180),
                    DialogResult = DialogResult.OK
                };
                okButton.Click += (s, args) =>
                {
                    selectedDate = calendar.SelectionStart;
                };

                if (calendarForm != null)
                {
                    calendarForm.Controls.Add(calendar);
                    calendarForm.Controls.Add(okButton);

                    if (calendarForm.ShowDialog() == DialogResult.OK)
                    {
                        ShowSuccess($"Selected date: {selectedDate.ToString("MMMM yyyy")}");
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error selecting date: {ex.Message}");
            }
        }

        private void BtnUploadTemplate_Click(object? sender, EventArgs e)
        {
            try
            {
                templateData = LoadExcelFile("Select Template File");
                if (templateData != null)
                {
                    ShowSuccess($"Template loaded: {templateData.Count} rows");
                    DisplayData(templateData, "Template Data");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading template: {ex.Message}");
            }
        }

        private void BtnUploadTimesheet_Click(object? sender, EventArgs e)
        {
            try
            {
                timesheetData = LoadExcelFile("Select Timesheet File");
                if (timesheetData != null)
                {
                    ShowSuccess($"Timesheet loaded: {timesheetData.Count} rows");
                    DisplayData(timesheetData, "Timesheet Data");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error loading timesheet: {ex.Message}");
            }
        }

        private void BtnProcess_Click(object? sender, EventArgs e)
        {
            try
            {
                if (templateData == null && timesheetData == null)
                {
                    ShowError("Please load at least one Excel file first.");
                    return;
                }

                string message = "Processing complete.\n";
                message += $"Selected Date: {selectedDate.ToString("MMMM yyyy")}\n";
                message += $"Template Data: {(templateData?.Count ?? 0)} rows\n";
                message += $"Timesheet Data: {(timesheetData?.Count ?? 0)} rows";

                ShowSuccess(message);
            }
            catch (Exception ex)
            {
                ShowError($"Processing error: {ex.Message}");
            }
        }

        private List<List<string>>? LoadExcelFile(string dialogTitle)
        {
            try
            {
                using OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = dialogTitle;
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return ParseExcelFile(openFileDialog.FileName);
                }
                return null;
            }
            catch (Exception ex)
            {
                ShowError($"Error loading file: {ex.Message}");
                return null;
            }
        }

        private List<List<string>>? ParseExcelFile(string filePath)
        {
            var data = new List<List<string>>();

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        ShowError("No worksheets found in the Excel file.");
                        return null;
                    }

                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int colCount = worksheet.Dimension?.Columns ?? 0;

                    if (rowCount == 0 || colCount == 0)
                    {
                        ShowError("The Excel file appears to be empty.");
                        return null;
                    }

                    for (int row = 1; row <= rowCount; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value;
                            rowData.Add(cellValue?.ToString() ?? string.Empty);
                        }
                        data.Add(rowData);
                    }

                    ShowSuccess($"Successfully parsed Excel file: {data.Count} rows, {colCount} columns");
                }
            }
            catch (Exception ex)
            {
                ShowError($"Error parsing Excel file: {ex.Message}");
                return null;
            }

            return data;
        }

        private void DisplayData(List<List<string>>? data, string title)
        {
            try
            {
                if (data == null || data.Count == 0)
                {
                    dataGrid?.Rows.Clear();
                    return;
                }
                dataGrid?.Columns.Clear();
                dataGrid?.Rows.Clear();

                // Create columns based on the first row
                int maxColumns = data.Max(row => row.Count);
                for (int i = 0; i < maxColumns; i++)
                {
                    dataGrid?.Columns.Add($"Column{i + 1}", $"Column {i + 1}");
                }

                // Add rows
                foreach (var row in data)
                {
                    var values = new object[maxColumns];
                    for (int i = 0; i < maxColumns; i++)
                    {
                        values[i] = i < row.Count ? row[i] : string.Empty;
                    }
                    dataGrid?.Rows.Add(values);
                }

                // Update the form title to show current data
                this.Text = $"Excel File Manager - {title}";
            }
            catch (Exception ex)
            {
                ShowError($"Error displaying data: {ex.Message}");
            }
        }

        private void ShowError(string? message)
        {
            if (message != null && errorLabel != null)
            {
                errorLabel.Text = message;
                errorLabel.ForeColor = Color.Red;
            }
        }

        private void ShowSuccess(string? message)
        {
            if (message != null && errorLabel != null)
            {
                errorLabel.Text = message;
                errorLabel.ForeColor = Color.Green;
            }
        }
    }
}