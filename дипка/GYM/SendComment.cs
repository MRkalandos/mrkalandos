using System;
using System.IO;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Net.NetworkInformation;
using System.Net.Mail;
using System.Net;
using MetroFramework;

namespace GYM
{
    public partial class SendComment : MetroForm
    {
        private const string TitleException = "Ошибка";
        public static bool attachment = false;

        public SendComment()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void Send_Mail_Load(object sender, EventArgs e)
        {
            FocusMe();
            try
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
                    HelperLog.Write(exception.Message);
                }

                if (internetStatus != IPStatus.Success)
                {
                    MetroMessageBox.Show(this, "Интернет-соединение отсутсвует, сообщение не отправиться",
                        "Соединение не установлено", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    FocusMe();
                }
                else
                {
                    MetroMessageBox.Show(this, "Интернет-соединение установлено, сообщение отправиться",
                        "Соединение установлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, "Сообщение не отправлено на почту" + metroTextBox4.Text, "Не отправлено",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                FocusMe();
                HelperLog.Write(exception.Message);
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                attachment = true;
                var dialog = new OpenFileDialog();
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    attachment = false;
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                FocusMe();
                HelperLog.Write(exception.Message);
            }
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "Вы уверены что хотите выйти?", "Подтверждение выхода",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            try
            {
                if
                ((metroTextBox1.Text == "") ||
                 (metroTextBox2.Text == "") ||
                 (metroTextBox3.Text == "") ||
                 (metroTextBox4.Text == "") ||
                 (metroTextBox5.Text == "") ||
                 (richTextBox1.Text == ""))
                {
                    MetroMessageBox.Show(this, "Заполните все поля", "Корректность ввода", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    FocusMe();
                }
                else
                {
                    var ping = new Ping();
                    var internetStatus = IPStatus.TimedOut;
                    var testSite = ping.Send(@"mail.ru");
                    if (testSite != null) internetStatus = testSite.Status;
                    if (internetStatus != IPStatus.Success)
                    {
                        MetroMessageBox.Show(this, "Интернет-соединение отсутсвует, сообщение не отправиться",
                            "Соединение не установлено", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        FocusMe();
                    }
                    else
                    {
                        var myAddress = new MailAddress(metroTextBox1.Text, metroTextBox3.Text);
                        var toAddress = new MailAddress(metroTextBox4.Text);
                        var objMessage = new MailMessage(myAddress, toAddress)
                        {
                            Subject = metroTextBox5.Text, Body = richTextBox1.Text
                        };
                        if (attachment)
                        {
                            attachment = false;
                            objMessage.Attachments.Add(new Attachment(null));
                        }

                        objMessage.IsBodyHtml = true;
                        var smtpServer = new SmtpClient("smtp.mail.ru", 587)
                        {
                            Credentials = new NetworkCredential(metroTextBox1.Text, metroTextBox2.Text),
                            EnableSsl = true
                        };
                        smtpServer.Send(objMessage);
                        MetroMessageBox.Show(this, "Сообщение отправлено на почту " + metroTextBox4.Text, "Отправлено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        FocusMe();
                    }
                }
            }
            catch (Exception exception)
            {
                HelperLog.Write(exception.Message);
            }
        }

        private void SendComment_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void SendComment_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5:
                    metroTile1.PerformClick();
                    FocusMe();
                    break;
                case Keys.Enter:
                    metroTile2.PerformClick();
                    FocusMe();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    FocusMe();
                    break;
                case Keys.Escape:
                    metroTile3.PerformClick();
                    FocusMe();
                    break;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    Help.ShowHelp(null, "Help/Help.chm");
                    FocusMe();
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
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

        private void SendComment_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
