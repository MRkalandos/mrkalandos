namespace GYM
{
    partial class Report_Employee
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.DSREPemployee = new GYM.DSREPemployee();
            this.vie1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.vie1TableAdapter = new GYM.DSREPemployeeTableAdapters.vie1TableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.DSREPemployee)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.vie1BindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSREPORTemployee";
            reportDataSource1.Value = this.vie1BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.Diagramm_Money.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 708);
            this.reportViewer1.TabIndex = 0;
            // 
            // DSREPemployee
            // 
            this.DSREPemployee.DataSetName = "DSREPemployee";
            this.DSREPemployee.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // vie1BindingSource
            // 
            this.vie1BindingSource.DataMember = "vie1";
            this.vie1BindingSource.DataSource = this.DSREPemployee;
            // 
            // vie1TableAdapter
            // 
            this.vie1TableAdapter.ClearBeforeFill = true;
            // 
            // Report_Employee
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 788);
            this.Controls.Add(this.reportViewer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Report_Employee";
            this.Text = "Диаграмма зарплат";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Report_Employee_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DSREPemployee)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.vie1BindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource vie1BindingSource;
        private DSREPemployee DSREPemployee;
        private DSREPemployeeTableAdapters.vie1TableAdapter vie1TableAdapter;
    }
}