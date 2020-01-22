namespace GYM
{
    partial class ReportTrenerovka
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
            this.DSRepVIE3 = new GYM.DSRepVIE3();
            this.vie3BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.vie3TableAdapter = new GYM.DSRepVIE3TableAdapters.vie3TableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.DSRepVIE3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.vie3BindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSrep";
            reportDataSource1.Value = this.vie3BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.REPORTTrenerovka.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(244, 185);
            this.reportViewer1.TabIndex = 0;
            // 
            // DSRepVIE3
            // 
            this.DSRepVIE3.DataSetName = "DSRepVIE3";
            this.DSRepVIE3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // vie3BindingSource
            // 
            this.vie3BindingSource.DataMember = "vie3";
            this.vie3BindingSource.DataSource = this.DSRepVIE3;
            // 
            // vie3TableAdapter
            // 
            this.vie3TableAdapter.ClearBeforeFill = true;
            // 
            // ReportTrenerovka
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.reportViewer1);
            this.Name = "ReportTrenerovka";
            this.Text = "Отчет по тренировкам";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.ReportTrenerovka_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DSRepVIE3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.vie3BindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource vie3BindingSource;
        private DSRepVIE3 DSRepVIE3;
        private DSRepVIE3TableAdapters.vie3TableAdapter vie3TableAdapter;
    }
}