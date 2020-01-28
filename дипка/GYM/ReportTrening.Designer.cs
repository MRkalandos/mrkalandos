namespace GYM
{
    partial class ReportTrening
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReportTrening));
            this.vie3BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DSRepVIE3 = new GYM.DSRepVIE3();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.vie3TableAdapter = new GYM.DSRepVIE3TableAdapters.vie3TableAdapter();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.metroToolTip1 = new MetroFramework.Components.MetroToolTip();
            ((System.ComponentModel.ISupportInitialize)(this.vie3BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSRepVIE3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // vie3BindingSource
            // 
            this.vie3BindingSource.DataMember = "vie3";
            this.vie3BindingSource.DataSource = this.DSRepVIE3;
            // 
            // DSRepVIE3
            // 
            this.DSRepVIE3.DataSetName = "DSRepVIE3";
            this.DSRepVIE3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
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
            this.reportViewer1.Click += new System.EventHandler(this.reportViewer1_Click);
            // 
            // vie3TableAdapter
            // 
            this.vie3TableAdapter.ClearBeforeFill = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(262, 243);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(23, 23);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            this.metroToolTip1.SetToolTip(this.pictureBox1, "Горячая клаивша F1");
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // metroToolTip1
            // 
            this.metroToolTip1.Style = MetroFramework.MetroColorStyle.Blue;
            this.metroToolTip1.StyleManager = null;
            this.metroToolTip1.Theme = MetroFramework.MetroThemeStyle.Light;
            // 
            // ReportTrening
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 265);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.reportViewer1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ReportTrening";
            this.Text = "Отчет по тренировкам";
            this.metroToolTip1.SetToolTip(this, "Горячая клавиша Esc");
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.ReportTrening_Activated);
            this.Load += new System.EventHandler(this.ReportTrenerovka_Load);
            this.Shown += new System.EventHandler(this.ReportTrening_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ReportTrening_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.vie3BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DSRepVIE3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource vie3BindingSource;
        private DSRepVIE3 DSRepVIE3;
        private DSRepVIE3TableAdapters.vie3TableAdapter vie3TableAdapter;
        private System.Windows.Forms.PictureBox pictureBox1;
        private MetroFramework.Components.MetroToolTip metroToolTip1;
    }
}