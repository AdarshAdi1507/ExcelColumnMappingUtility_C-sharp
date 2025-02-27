namespace ExcelProcessor.Forms
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.TextBox txtXmlPath;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Button btnBrowseXml;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ComboBox cmbSheets;
        private System.Windows.Forms.Label lblSheet;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Panel panelMain;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panelMain = new System.Windows.Forms.Panel();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.txtXmlPath = new System.Windows.Forms.TextBox();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.btnBrowseXml = new System.Windows.Forms.Button();
            this.btnProcess = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.cmbSheets = new System.Windows.Forms.ComboBox();
            this.lblSheet = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();

            // panelMain
            this.panelMain.Padding = new System.Windows.Forms.Padding(20);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.BackColor = System.Drawing.SystemColors.Window;

            // txtExcelPath
            this.txtExcelPath.Location = new System.Drawing.Point(20, 20);
            this.txtExcelPath.Size = new System.Drawing.Size(400, 23);
            this.txtExcelPath.ReadOnly = true;
            this.txtExcelPath.BackColor = System.Drawing.SystemColors.Window;

            // btnBrowseExcel
            this.btnBrowseExcel.Location = new System.Drawing.Point(430, 19);
            this.btnBrowseExcel.Size = new System.Drawing.Size(100, 25);
            this.btnBrowseExcel.Text = "Browse Excel";
            this.btnBrowseExcel.UseVisualStyleBackColor = true;
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);

            // lblSheet
            this.lblSheet.Location = new System.Drawing.Point(20, 53);
            this.lblSheet.Size = new System.Drawing.Size(80, 23);
            this.lblSheet.Text = "Select Sheet:";
            this.lblSheet.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            // cmbSheets
            this.cmbSheets.Location = new System.Drawing.Point(100, 53);
            this.cmbSheets.Size = new System.Drawing.Size(320, 23);
            this.cmbSheets.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSheets.SelectedIndexChanged += new System.EventHandler(this.cmbSheets_SelectedIndexChanged);

            // txtXmlPath
            this.txtXmlPath.Location = new System.Drawing.Point(20, 86);
            this.txtXmlPath.Size = new System.Drawing.Size(400, 23);
            this.txtXmlPath.ReadOnly = true;
            this.txtXmlPath.BackColor = System.Drawing.SystemColors.Window;

            // btnBrowseXml
            this.btnBrowseXml.Location = new System.Drawing.Point(430, 85);
            this.btnBrowseXml.Size = new System.Drawing.Size(100, 25);
            this.btnBrowseXml.Text = "Browse XML";
            this.btnBrowseXml.UseVisualStyleBackColor = true;
            this.btnBrowseXml.Click += new System.EventHandler(this.btnBrowseXml_Click);

            // btnProcess
            this.btnProcess.Location = new System.Drawing.Point(540, 19);
            this.btnProcess.Size = new System.Drawing.Size(100, 91);
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);

            // progressBar
            this.progressBar.Location = new System.Drawing.Point(20, 119);
            this.progressBar.Size = new System.Drawing.Size(620, 23);
            this.progressBar.Visible = false;

            // lblStatus
            this.lblStatus.Location = new System.Drawing.Point(20, 152);
            this.lblStatus.AutoSize = true;
            this.lblStatus.ForeColor = System.Drawing.Color.DarkBlue;

            // Form1
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(664, 191);
            this.MinimumSize = new System.Drawing.Size(680, 230);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.cmbSheets);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnBrowseXml);
            this.Controls.Add(this.txtXmlPath);
            this.Controls.Add(this.btnBrowseExcel);
            this.Controls.Add(this.txtExcelPath);
            this.Controls.Add(this.panelMain);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Processor";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}