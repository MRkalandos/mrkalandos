using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ReportTrening : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                   Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
                                   "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/ReportTrening.txt";

        public ReportTrening()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void ReportTrenerovka_Load(object sender, EventArgs e)
        {
            try
            {
                FocusMe();
                var connection = new OleDbConnection(conString);
                connection.Open();
                var adapter = new OleDbDataAdapter(@"select * from vie3", connection);
                adapter.Fill(DSRepVIE3.vie3);
                connection.Close();
                this.reportViewer1.RefreshReport();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, "Отчет не сформирован", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    FocusMe();
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void ReportTrening_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ReportTrening_KeyDown(object sender, KeyEventArgs e)
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                FocusMe();
                if (File.Exists("Help/ReportTrening.chm"))
                {
                    Help.ShowHelp(null, "Help/ReportTrening.chm");
                    FocusMe();
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException,MessageBoxButtons.OK,MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
                else
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (sw.BaseStream).Seek(0, SeekOrigin.End);
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
            }
        }

        private void reportViewer1_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ReportTrening_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}

