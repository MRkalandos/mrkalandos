using System;
using System.IO;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class EmailRePort : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/EmailReport.txt";

        public EmailRePort()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void EmailREPort_Load(object sender, EventArgs e)
        {
            var internetStatus = IPStatus.TimedOut;
            try
            {
                var ping = new Ping();
                var testSite = ping.Send(@"yandex.ru");
                if (testSite != null) internetStatus = testSite.Status;

            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException);
            }

            if (internetStatus != IPStatus.Success)
            {
                MetroMessageBox.Show(this, "Интернет-соединение отсутсвует, сообщение не отправиться",
                    "Соединение не установлено", MessageBoxButtons.OK);
                FocusMe();
            }
            else
            {
                MetroMessageBox.Show(this, "Интернет-соединение установлено, сообщение отправиться",
                    "Соединение установлено", MessageBoxButtons.OK);
                FocusMe();
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if ((metroTextBox4.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/EmailReport.chm"))
                {
                    Help.ShowHelp(null, "Help/EmailReport.chm");
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException,MessageBoxButtons.OK,MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void EmailRePort_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void EmailRePort_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void EmailRePort_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5:
                    metroTile2.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    Close();
                    break;
            }
        }
    }
}
