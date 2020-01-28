namespace GYM
{
    partial class Exit
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
            this.metroTile4 = new MetroFramework.Controls.MetroTile();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.metroTile6 = new MetroFramework.Controls.MetroTile();
            this.metroTile5 = new MetroFramework.Controls.MetroTile();
            this.metroLabel31 = new MetroFramework.Controls.MetroLabel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.metroTile4.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroTile4
            // 
            this.metroTile4.ActiveControl = null;
            this.metroTile4.Controls.Add(this.metroLabel1);
            this.metroTile4.Controls.Add(this.metroTile6);
            this.metroTile4.Controls.Add(this.metroTile5);
            this.metroTile4.Controls.Add(this.metroLabel31);
            this.metroTile4.Location = new System.Drawing.Point(-551, -263);
            this.metroTile4.Name = "metroTile4";
            this.metroTile4.Size = new System.Drawing.Size(2191, 1057);
            this.metroTile4.TabIndex = 13;
            this.metroTile4.Text = "metroTile4";
            this.metroTile4.UseSelectable = true;
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel1.Location = new System.Drawing.Point(1194, 510);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(0, 0);
            this.metroLabel1.TabIndex = 3;
            // 
            // metroTile6
            // 
            this.metroTile6.ActiveControl = null;
            this.metroTile6.Location = new System.Drawing.Point(1311, 641);
            this.metroTile6.Name = "metroTile6";
            this.metroTile6.Size = new System.Drawing.Size(75, 44);
            this.metroTile6.Style = MetroFramework.MetroColorStyle.Red;
            this.metroTile6.TabIndex = 2;
            this.metroTile6.Text = "Нет";
            this.metroTile6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.metroTile6.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.metroTile6.UseSelectable = true;
            this.metroTile6.Click += new System.EventHandler(this.metroTile6_Click);
            // 
            // metroTile5
            // 
            this.metroTile5.ActiveControl = null;
            this.metroTile5.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.metroTile5.Location = new System.Drawing.Point(1137, 641);
            this.metroTile5.Name = "metroTile5";
            this.metroTile5.Size = new System.Drawing.Size(75, 44);
            this.metroTile5.Style = MetroFramework.MetroColorStyle.Red;
            this.metroTile5.TabIndex = 1;
            this.metroTile5.Text = "Да";
            this.metroTile5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.metroTile5.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.metroTile5.UseSelectable = true;
            this.metroTile5.Click += new System.EventHandler(this.metroTile5_Click);
            // 
            // metroLabel31
            // 
            this.metroLabel31.AutoSize = true;
            this.metroLabel31.FontSize = MetroFramework.MetroLabelSize.Tall;
            this.metroLabel31.Location = new System.Drawing.Point(1081, 598);
            this.metroLabel31.Name = "metroLabel31";
            this.metroLabel31.Size = new System.Drawing.Size(371, 25);
            this.metroLabel31.TabIndex = 0;
            this.metroLabel31.Text = "Вы уверены что хотите выйти из программы?";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Exit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1386, 788);
            this.Controls.Add(this.metroTile4);
            this.Name = "Exit";
            this.Text = "Exit";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Exit_Load);
            this.metroTile4.ResumeLayout(false);
            this.metroTile4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroTile metroTile4;
        private MetroFramework.Controls.MetroLabel metroLabel31;
        private MetroFramework.Controls.MetroTile metroTile6;
        public MetroFramework.Controls.MetroTile metroTile5;
        public MetroFramework.Controls.MetroLabel metroLabel1;
        private System.Windows.Forms.Timer timer1;
    }
}