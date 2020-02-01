using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ReportSportsmen : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                   Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
                                   "/ISgym.mdb;Jet OLEDB:Database Password=316206");

      public ReportSportsmen()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void RepSportsmen_Load(object sender, EventArgs e)
        {
            try
            {
                FocusMe();
                var connection = new OleDbConnection(conString);
                connection.Open();
                var adapterReportSportsmen = new OleDbDataAdapter(@"select * from vie2", connection);
                adapterReportSportsmen.Fill(DSREPSportsmenVIE2.vie2);
                connection.Close();
                this.reportViewer1.RefreshReport();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    FocusMe();
                    Help.ShowHelp(null, "Help/Help.chm");
                    FocusMe();
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
        }

        private void ReportSportsmen_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Close();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    FocusMe();
                    break;
            }
        }

        private void ReportSportsmen_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ReportSportsmen_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}