using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ReportTrener : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public ReportTrener()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void Report_Trener_Load(object sender, EventArgs e)
        {
            try
            {
                var connection = new OleDbConnection(ConnectionDb.conString());
                connection.Open();
                var adapter = new OleDbDataAdapter(@"select * from тренер", connection);
                adapter.Fill(RepTrener.Тренер);
                connection.Close();
                this.reportViewer1.RefreshReport();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.ToString());
            }
            finally
            {
                FocusMe();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    Help.ShowHelp(null, "Help/Help.chm");
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.ToString());
            }
            finally
            {
                FocusMe();
            }
        }

        private void ReportTrener_KeyDown(object sender, KeyEventArgs e)
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

        private void ReportTrener_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void reportViewer1_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ReportTrener_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ReportTrener_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
