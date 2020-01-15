namespace GYM
{
    partial class Report
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
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource2 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.Зарплата_сотрудникаBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DS_Money = new GYM.DS_Money();
            this.СотрудникBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DSEmployee = new GYM.DSEmployee();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.Зарплата_сотрудникаTableAdapter = new GYM.DS_MoneyTableAdapters.Зарплата_сотрудникаTableAdapter();
            this.СотрудникTableAdapter = new GYM.DSEmployeeTableAdapters.СотрудникTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.Зарплата_сотрудникаBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DS_Money)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.СотрудникBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSEmployee)).BeginInit();
            this.SuspendLayout();
            // 
            // Зарплата_сотрудникаBindingSource
            // 
            this.Зарплата_сотрудникаBindingSource.DataMember = "Зарплата_сотрудника";
            this.Зарплата_сотрудникаBindingSource.DataSource = this.DS_Money;
            // 
            // DS_Money
            // 
            this.DS_Money.DataSetName = "DS_Money";
            this.DS_Money.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // СотрудникBindingSource
            // 
            this.СотрудникBindingSource.DataMember = "Сотрудник";
            this.СотрудникBindingSource.DataSource = this.DSEmployee;
            // 
            // DSEmployee
            // 
            this.DSEmployee.DataSetName = "DSEmployee";
            this.DSEmployee.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.Зарплата_сотрудникаBindingSource;
            reportDataSource2.Name = "DataSet2";
            reportDataSource2.Value = this.СотрудникBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 0);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(602, 265);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Visible = false;
            // 
            // Зарплата_сотрудникаTableAdapter
            // 
            this.Зарплата_сотрудникаTableAdapter.ClearBeforeFill = true;
            // 
            // СотрудникTableAdapter
            // 
            this.СотрудникTableAdapter.ClearBeforeFill = true;
            // 
            // Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 265);
            this.Controls.Add(this.reportViewer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Report";
            this.Text = "Report";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Report_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Зарплата_сотрудникаBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DS_Money)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.СотрудникBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSEmployee)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.BindingSource Зарплата_сотрудникаBindingSource;
        private DS_Money DS_Money;
        private System.Windows.Forms.BindingSource СотрудникBindingSource;
        private DSEmployee DSEmployee;
        private DS_MoneyTableAdapters.Зарплата_сотрудникаTableAdapter Зарплата_сотрудникаTableAdapter;
        private DSEmployeeTableAdapters.СотрудникTableAdapter СотрудникTableAdapter;
        public Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
    }
}