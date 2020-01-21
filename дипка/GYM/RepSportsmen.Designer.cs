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
            this.СпортсменBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DsRepSportsmen = new GYM.DsRepSportsmen();
            this.СпортсменTableAdapter = new GYM.DsRepSportsmenTableAdapters.СпортсменTableAdapter();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            ((System.ComponentModel.ISupportInitialize)(this.СпортсменBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DsRepSportsmen)).BeginInit();
            this.SuspendLayout();
            // 
            // СпортсменBindingSource
            // 
            this.СпортсменBindingSource.DataMember = "Спортсмен";
            this.СпортсменBindingSource.DataSource = this.DsRepSportsmen;
            // 
            // DsRepSportsmen
            // 
            this.DsRepSportsmen.DataSetName = "DsRepSportsmen";
            this.DsRepSportsmen.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // СпортсменTableAdapter
            // 
            this.СпортсменTableAdapter.ClearBeforeFill = true;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportViewer1.Location = new System.Drawing.Point(20, 60);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(1346, 708);
            this.reportViewer1.TabIndex = 0;
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
            this.Text = "Отчет по посещениям";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.RepSportsmen_Load);
            ((System.ComponentModel.ISupportInitialize)(this.СпортсменBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DsRepSportsmen)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.BindingSource СпортсменBindingSource;
        private DsRepSportsmen DsRepSportsmen;
        private DsRepSportsmenTableAdapters.СпортсменTableAdapter СпортсменTableAdapter;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
    }
}