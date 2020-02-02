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
                var testSite = ping.Send(@"mail.ru");
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
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
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

        private void EmailRePort_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
