namespace GYM
{
    partial class ReportSportsmen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportSportsmen));
            this.vie2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DSREPSportsmenVIE2 = new GYM.DSREPSportsmenVIE2();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.vie2TableAdapter = new GYM.DSREPSportsmenVIE2TableAdapters.vie2TableAdapter();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.metroToolTip1 = new MetroFramework.Components.MetroToolTip();
            ((System.ComponentModel.ISupportInitialize)(this.vie2BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSREPSportsmenVIE2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
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
            reportDataSource2.Name = "DSRepSportsmen";
            reportDataSource2.Value = this.vie2BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "GYM.RepSportsmen.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 708);
            this.reportViewer1.TabIndex = 0;
            this.reportViewer1.Click += new System.EventHandler(this.reportViewer1_Click);
            // 
            // vie2TableAdapter
            // 
            this.vie2TableAdapter.ClearBeforeFill = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(1365, 766);
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
            // ReportSportsmen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 788);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.reportViewer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReportSportsmen";
            this.Text = "Отчет по посещениям спортсменов";
            this.metroToolTip1.SetToolTip(this, "Esc");
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.ReportSportsmen_Activated);
            this.Load += new System.EventHandler(this.RepSportsmen_Load);
            this.Shown += new System.EventHandler(this.ReportSportsmen_Shown);
            this.Click += new System.EventHandler(this.ReportSportsmen_Click);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ReportSportsmen_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.vie2BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSREPSportsmenVIE2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource vie2BindingSource;
        private DSREPSportsmenVIE2 DSREPSportsmenVIE2;
        private DSREPSportsmenVIE2TableAdapters.vie2TableAdapter vie2TableAdapter;
        private System.Windows.Forms.PictureBox pictureBox1;
        private MetroFramework.Components.MetroToolTip metroToolTip1;
    }
}