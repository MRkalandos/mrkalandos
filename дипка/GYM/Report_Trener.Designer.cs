namespace GYM
{
    partial class Report_Trener
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
            this.ТренерBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.RepTrener = new GYM.RepTrener();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.ТренерTableAdapter = new GYM.RepTrenerTableAdapters.ТренерTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.ТренерBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepTrener)).BeginInit();
            this.SuspendLayout();
            // 
            // ТренерBindingSource
            // 
            this.ТренерBindingSource.DataMember = "Тренер";
            this.ТренерBindingSource.DataSource = this.RepTrener;
            // 
            // RepTrener
            // 
            this.RepTrener.DataSetName = "RepTrener";
            this.RepTrener.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSREPTRENER";
            reportDataSource1.Value = this.ТренерBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.DiagrammMoneyTrener.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 708);
            this.reportViewer1.TabIndex = 0;
            // 
            // ТренерTableAdapter
            // 
            this.ТренерTableAdapter.ClearBeforeFill = true;
            // 
            // Report_Trener
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 788);
            this.Controls.Add(this.reportViewer1);
            this.Name = "Report_Trener";
            this.Text = "Диаграмма оклада тренеров";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Report_Trener_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ТренерBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepTrener)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource ТренерBindingSource;
        private RepTrener RepTrener;
        private RepTrenerTableAdapters.ТренерTableAdapter ТренерTableAdapter;
    }
}