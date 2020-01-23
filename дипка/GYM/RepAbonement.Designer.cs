namespace GYM
{
    partial class RepAbone_ent
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
            this.REPAbonement = new GYM.REPAbonement();
            this.АбонементBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.АбонементTableAdapter = new GYM.REPAbonementTableAdapters.АбонементTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.REPAbonement)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.АбонементBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSrepAbonement";
            reportDataSource1.Value = this.АбонементBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.REPAbonement.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(244, 185);
            this.reportViewer1.TabIndex = 0;
            // 
            // REPAbonement
            // 
            this.REPAbonement.DataSetName = "REPAbonement";
            this.REPAbonement.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // АбонементBindingSource
            // 
            this.АбонементBindingSource.DataMember = "Абонемент";
            this.АбонементBindingSource.DataSource = this.REPAbonement;
            // 
            // АбонементTableAdapter
            // 
            this.АбонементTableAdapter.ClearBeforeFill = true;
            // 
            // RepAbone_ent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.reportViewer1);
            this.Name = "RepAbone_ent";
            this.Text = "Отчет по абонементам";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.RepAbone_ent_Load);
            ((System.ComponentModel.ISupportInitialize)(this.REPAbonement)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.АбонементBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource АбонементBindingSource;
        private REPAbonement REPAbonement;
        private REPAbonementTableAdapters.АбонементTableAdapter АбонементTableAdapter;
    }
}