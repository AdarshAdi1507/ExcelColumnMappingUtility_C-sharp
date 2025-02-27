// Form1.cs - Replace the entire content of your existing Form1.cs
using System;
using System.Windows.Forms;
using System.IO;
using ExcelProcessor.Interfaces;
using ExcelProcessor.Services;
using System.Drawing;
using OfficeOpenXml;
using System.Linq;

namespace ExcelProcessor.Forms
{
    public partial class Form1 : Form
    {
        private readonly IExcelService _excelService;
        private readonly ILogService _logService;
        private readonly IXmlMappingService _xmlMappingService;

        public Form1()
        {
            InitializeComponent();
            _logService = new LogService();
            _excelService = new ExcelService(_logService);
            _xmlMappingService = new XmlMappingService(_logService);

            // Set initial state of controls
            btnProcess.Enabled = false;
            cmbSheets.Enabled = false;
        }

        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select an Excel File";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelPath.Text = openFileDialog.FileName;
                    LoadExcelSheets(openFileDialog.FileName);
                }
            }
        }

        private void LoadExcelSheets(string filePath)
        {
            try
            {
                cmbSheets.Items.Clear();
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheets = package.Workbook.Worksheets.Select(s => s.Name).ToArray();
                    cmbSheets.Items.AddRange(sheets);
                    if (sheets.Length > 0)
                    {
                        cmbSheets.SelectedIndex = 0;
                        cmbSheets.Enabled = true;
                    }
                }
                ValidateInputs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading Excel sheets: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbSheets.Enabled = false;
            }
        }

        private void btnBrowseXml_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "XML Files|*.xml";
                openFileDialog.Title = "Select Mapping XML File";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtXmlPath.Text = openFileDialog.FileName;
                    ValidateInputs();
                }
            }
        }

        private void ValidateInputs()
        {
            btnProcess.Enabled = !string.IsNullOrEmpty(txtExcelPath.Text) &&
                               !string.IsNullOrEmpty(txtXmlPath.Text) &&
                               cmbSheets.SelectedIndex >= 0;
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtExcelPath.Text) || string.IsNullOrEmpty(txtXmlPath.Text))
            {
                MessageBox.Show("Please select both Excel and XML mapping files.", "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Disable controls during processing
                btnProcess.Enabled = false;
                btnBrowseExcel.Enabled = false;
                btnBrowseXml.Enabled = false;
                cmbSheets.Enabled = false;
                progressBar.Value = 0;
                progressBar.Visible = true;

                // Update status
                lblStatus.Text = "Reading configuration...";
                progressBar.Value = 20;

                // Read mapping configuration
                var (mappings, headerRow, startRow) = _xmlMappingService.ReadMappingConfiguration(txtXmlPath.Text);

                lblStatus.Text = "Validating source file...";
                progressBar.Value = 40;

                // Get total columns from source file
                int totalColumns = _excelService.GetTotalColumns(txtExcelPath.Text, cmbSheets.SelectedItem.ToString());

                // Validate configuration
                _xmlMappingService.ValidateConfiguration(mappings, headerRow, startRow, totalColumns);

                lblStatus.Text = "Processing Excel file...";
                progressBar.Value = 60;

                // Process and generate new Excel
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(txtExcelPath.Text),
                    $"processed_{Path.GetFileNameWithoutExtension(txtExcelPath.Text)}.xlsx"
                );

                _excelService.ProcessAndGenerateTeamcenterExcel(
                    txtExcelPath.Text,
                    mappings,
                    outputPath,
                    headerRow,
                    startRow,
                    cmbSheets.SelectedItem.ToString()
                );

                progressBar.Value = 100;
                lblStatus.Text = $"Successfully generated Excel file: {Path.GetFileName(outputPath)}";
                _logService.LogInformation($"Successfully generated Excel file: {outputPath}");

                MessageBox.Show("Excel file generated successfully!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing files: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _logService.LogError("Error processing files", ex);
            }
            finally
            {
                // Re-enable controls
                btnProcess.Enabled = true;
                btnBrowseExcel.Enabled = true;
                btnBrowseXml.Enabled = true;
                cmbSheets.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void cmbSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            ValidateInputs();
        }
    }
}