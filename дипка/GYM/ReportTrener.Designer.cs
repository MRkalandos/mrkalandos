namespace GYM
{
    partial class ReportTrener
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
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource2 = new Microsoft.Reporting.WinForms.ReportDataSource();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportTrener));
            this.ТренерBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.RepTrener = new GYM.RepTrener();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.ТренерTableAdapter = new GYM.RepTrenerTableAdapters.ТренерTableAdapter();
            this.metroToolTip1 = new MetroFramework.Components.MetroToolTip();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.ТренерBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepTrener)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
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
            reportDataSource2.Name = "DSREPTRENER";
            reportDataSource2.Value = this.ТренерBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.DiagrammMoneyTrener.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 693);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Click += new System.EventHandler(this.reportViewer1_Click);
            // 
            // ТренерTableAdapter
            // 
            this.ТренерTableAdapter.ClearBeforeFill = true;
            // 
            // metroToolTip1
            // 
            this.metroToolTip1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroToolTip1.StyleManager = null;
            this.metroToolTip1.Theme = MetroFramework.MetroThemeStyle.Light;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(1364, 751);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(23, 23);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // ReportTrener
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 773);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.reportViewer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReportTrener";
            this.Text = "Диаграмма оклада тренеров";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.ReportTrener_Activated);
            this.Load += new System.EventHandler(this.Report_Trener_Load);
            this.Shown += new System.EventHandler(this.ReportTrener_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ReportTrener_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.ТренерBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepTrener)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource ТренерBindingSource;
        private RepTrener RepTrener;
        private RepTrenerTableAdapters.ТренерTableAdapter ТренерTableAdapter;
        private MetroFramework.Components.MetroToolTip metroToolTip1;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}