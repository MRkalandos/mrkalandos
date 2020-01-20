namespace GYM
{
    partial class RepSportsmen
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
            this.DsRepSportsmen = new GYM.DsRepSportsmen();
            this.СпортсменBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.СпортсменTableAdapter = new GYM.DsRepSportsmenTableAdapters.СпортсменTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.DsRepSportsmen)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.СпортсменBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSRepSportsmen";
            reportDataSource1.Value = this.СпортсменBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.Report3.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 0);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(284, 265);
            this.reportViewer1.TabIndex = 0;
            // 
            // DsRepSportsmen
            // 
            this.DsRepSportsmen.DataSetName = "DsRepSportsmen";
            this.DsRepSportsmen.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // СпортсменBindingSource
            // 
            this.СпортсменBindingSource.DataMember = "Спортсмен";
            this.СпортсменBindingSource.DataSource = this.DsRepSportsmen;
            // 
            // СпортсменTableAdapter
            // 
            this.СпортсменTableAdapter.ClearBeforeFill = true;
            // 
            // RepSportsmen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.reportViewer1);
            this.Name = "RepSportsmen";
            this.Text = "RepSportsmen";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.RepSportsmen_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DsRepSportsmen)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.СпортсменBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource СпортсменBindingSource;
        private DsRepSportsmen DsRepSportsmen;
        private DsRepSportsmenTableAdapters.СпортсменTableAdapter СпортсменTableAdapter;
    }
}