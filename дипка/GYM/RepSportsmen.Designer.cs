﻿namespace GYM
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
            this.vie2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DSREPSportsmenVIE2 = new GYM.DSREPSportsmenVIE2();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.vie2TableAdapter = new GYM.DSREPSportsmenVIE2TableAdapters.vie2TableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.vie2BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSREPSportsmenVIE2)).BeginInit();
            this.SuspendLayout();
            // 
            // vie2BindingSource
            // 
            this.vie2BindingSource.DataMember = "vie2";
            this.vie2BindingSource.DataSource = this.DSREPSportsmenVIE2;
            // 
            // DSREPSportsmenVIE2
            // 
            this.DSREPSportsmenVIE2.DataSetName = "DSREPSportsmenVIE2";
            this.DSREPSportsmenVIE2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DSRepSportsmen";
            reportDataSource1.Value = this.vie2BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.RepSportsmen.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 708);
            this.reportViewer1.TabIndex = 0;
            // 
            // vie2TableAdapter
            // 
            this.vie2TableAdapter.ClearBeforeFill = true;
            // 
            // RepSportsmen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 788);
            this.Controls.Add(this.reportViewer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RepSportsmen";
            this.Text = "Отчет по посещениям спортсменов";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.RepSportsmen_Load);
            ((System.ComponentModel.ISupportInitialize)(this.vie2BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSREPSportsmenVIE2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource vie2BindingSource;
        private DSREPSportsmenVIE2 DSREPSportsmenVIE2;
        private DSREPSportsmenVIE2TableAdapters.vie2TableAdapter vie2TableAdapter;
    }
}