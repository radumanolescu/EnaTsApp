using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using Com.Ena.Timesheet.Xl;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using Com.Ena.Timesheet.Phd;
using Com.Ena.Timesheet.Ena;
using Com.Ena.Timesheet;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using System.IO;

namespace EnaTsApp
{
    public class MainForm : Form
    {
        private readonly ILogger<MainForm> _logger;
        private readonly IServiceProvider _serviceProvider;

        private Button? btnChooseMonth;
        private Button? btnUploadTemplate;
        private Button? btnUploadTimesheet;
        private Button? btnProcess;
        private Panel? errorPanel;
        private Label? errorLabel;
        private DataGridView? dataGrid;        
        private Dictionary<string, string> fileLocations = new Dictionary<string, string>();

        private DateTime selectedDate = DateTime.Now;
        private List<List<string>>? templateData;
        private List<List<string>>? timesheetData;
        private PhdTemplate? phdTemplate;
        private EnaTimesheet? enaTimesheet;

        public MainForm(ILogger<MainForm> logger, IServiceProvider serviceProvider)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            this.Load += MainForm_Load;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                _logger.LogInformation("MainForm initialized");
                // Ensure logs directory exists
                var logsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EnaTsApp", "logs");
                Directory.CreateDirectory(logsDirectory);
                _logger.LogInformation($"Logs directory created: {logsDirectory}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error initializing MainForm");
                throw;
            }
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
            this.MinimumSize = new Size(1600, 900); // Set minimum size
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Font; // Enable font-based scaling

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
                Size = new Size(240, 30),
                Location = new Point(10, 10)
            };
            btnChooseMonth.Click += BtnChooseMonth_Click;

            // Button 2: Upload Template
            btnUploadTemplate = new Button
            {
                Text = "Upload Template",
                Size = new Size(240, 30),
                Location = new Point(260, 10)
            };
            btnUploadTemplate.Click += BtnUploadTemplate_Click;

            // Button 3: Upload Timesheet
            btnUploadTimesheet = new Button
            {
                Text = "Upload Timesheet",
                Size = new Size(240, 30),
                Location = new Point(510, 10)
            };
            btnUploadTimesheet.Click += BtnUploadTimesheet_Click;

            // Button 4: Process
            btnProcess = new Button
            {
                Text = "Process Data",
                Size = new Size(240, 30),
                Location = new Point(760, 10)
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
                BackColor = Color.FromArgb(255, 245, 245),
                AutoSize = true
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
                AutoSize = true,
                ForeColor = Color.Green,
                Location = new Point(10, 25)
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
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
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
                    Size = new Size(300, 300),
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
                    Size = new Size(75, 30),
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
                templateData = LoadExcelFile("Select Template File", "template");
                if (templateData != null)
                {
                    ShowSuccess($"Template loaded: {templateData.Count} rows");
                    DisplayData(templateData, "Template Data");
                    string templateFilePath = fileLocations["template"];
                    string outputFilePath = Path.Combine(Path.GetDirectoryName(templateFilePath) ?? throw new InvalidOperationException("Template file path is invalid"), $"PHD ENA Timesheet {selectedDate.ToString("yyyy-MM")}.xlsx");
                    phdTemplate = new PhdTemplate(selectedDate.ToString("yyyyMM"), templateData, templateFilePath, outputFilePath);
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
                try
                {
                    var timesheetLogger = _serviceProvider.GetRequiredService<ILogger<EnaTimesheet>>();
                    var entryLogger = _serviceProvider.GetRequiredService<ILogger<EnaTsEntry>>();
                    timesheetData = LoadExcelFile("Select Timesheet File", "timesheet");
                    if (timesheetData != null)
                    {
                        ShowSuccess($"Timesheet loaded: {timesheetData.Count} rows");
                        DisplayData(timesheetData, "Timesheet Data");
                        string timesheetFilePath = fileLocations["timesheet"];
                        string outputFilePath = Path.Combine(Path.GetDirectoryName(timesheetFilePath) ?? throw new InvalidOperationException("Timesheet file path is invalid"), $"PHD ENA Timesheet {selectedDate.ToString("yyyy-MM")}.xlsx");
                        enaTimesheet = new EnaTimesheet(selectedDate.ToString("yyyyMM"), timesheetData, timesheetFilePath, outputFilePath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error creating EnaTimesheet");
                    throw;
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
                if (templateData == null || timesheetData == null)
                {
                    ShowError("Please load a template and a timesheet first.");
                    return;
                }
                string yyyyMM = selectedDate.ToString("yyyyMM");
                string templatePath = fileLocations["template"];
                string timesheetPath = fileLocations["timesheet"];
                TimesheetProcessor processor = new TimesheetProcessor(yyyyMM, templatePath, timesheetPath);
                var invalidActivities = processor.Validate();
                var outputFile = processor.Process();

                string message = $"See output file: {outputFile}\n";
                message += $"Selected Date: {selectedDate.ToString("MMMM yyyy")}. ";
                message += $"Template Data: {(templateData?.Count ?? 0)} rows. ";
                message += $"Timesheet Data: {(timesheetData?.Count ?? 0)} rows. ";

                ShowSuccess(message);
            }
            catch (Exception ex)
            {
                ShowError($"Processing error: {ex.Message}");
            }
        }

        private List<List<string>>? LoadExcelFile(string dialogTitle, string fileKey)
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
                    fileLocations[fileKey] = openFileDialog.FileName;
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
            try
            {
                var parser = new ExcelParser();
                var data = parser.ParseExcelFile(filePath);
                if (data != null)
                {
                    ShowSuccess($"Successfully parsed Excel file: {data.Count} rows, {data[0].Count} columns");
                }
                return data;
            }
            catch (InvalidOperationException ex)
            {
                ShowError(ex.Message);
                return null;
            }
            catch (Exception ex)
            {
                ShowError($"Unexpected error parsing Excel file: {ex.Message}");
                return null;
            }
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