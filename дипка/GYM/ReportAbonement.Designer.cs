namespace GYM
{
    partial class ReportAbonement
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportAbonement));
            this.АбонементBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.REPAbonement = new GYM.REPAbonement();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.АбонементTableAdapter = new GYM.REPAbonementTableAdapters.АбонементTableAdapter();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.metroToolTip1 = new MetroFramework.Components.MetroToolTip();
            ((System.ComponentModel.ISupportInitialize)(this.АбонементBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.REPAbonement)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // АбонементBindingSource
            // 
            this.АбонементBindingSource.DataMember = "Абонемент";
            this.АбонементBindingSource.DataSource = this.REPAbonement;
            // 
            // REPAbonement
            // 
            this.REPAbonement.DataSetName = "REPAbonement";
            this.REPAbonement.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource2.Name = "DSrepAbonement";
            reportDataSource2.Value = this.АбонементBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.REPAbonement.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(244, 185);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Click += new System.EventHandler(this.reportViewer1_Click);
            // 
            // АбонементTableAdapter
            // 
            this.АбонементTableAdapter.ClearBeforeFill = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(263, 243);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(23, 23);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            this.metroToolTip1.SetToolTip(this.pictureBox1, "F1");
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // metroToolTip1
            // 
            this.metroToolTip1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroToolTip1.StyleManager = null;
            this.metroToolTip1.Theme = MetroFramework.MetroThemeStyle.Light;
            // 
            // ReportAbonement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.reportViewer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReportAbonement";
            this.Text = "Отчет по абонементам";
            this.metroToolTip1.SetToolTip(this, "Esc");
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.RepAbone_ent_Activated);
            this.Load += new System.EventHandler(this.RepAbone_ent_Load);
            this.Shown += new System.EventHandler(this.RepAbone_ent_Shown);
            this.Click += new System.EventHandler(this.ReportAbonement_Click);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RepAbone_ent_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.АбонементBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.REPAbonement)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource АбонементBindingSource;
        private REPAbonement REPAbonement;
        private REPAbonementTableAdapters.АбонементTableAdapter АбонементTableAdapter;
        private System.Windows.Forms.PictureBox pictureBox1;
        private MetroFramework.Components.MetroToolTip metroToolTip1;
    }
}