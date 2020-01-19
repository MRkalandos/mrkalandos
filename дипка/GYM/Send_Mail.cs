using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Net.NetworkInformation;
using System.Net.Mail;
using System.Net;
using System.IO;

namespace GYM
{
    public partial class Send_Mail : MetroForm
    {
        public Send_Mail()
        {
            InitializeComponent();
        }
        
        private void Send_Mail_Load(object sender, EventArgs e)
        {
            IPStatus status = IPStatus.TimedOut;
            try
            {
                Ping ping = new Ping();
                PingReply reply = ping.Send(@"yandex.ru");
                status = reply.Status;
            }
            catch { }
            if (status != IPStatus.Success)
            {
                MessageBox.Show("Интернет соединение отсутсвует, сообщение не отправиться","Соединение не установлено");
            }
            else
            {
                MessageBox.Show("Интернет-соединение установлено", "Соединение установлено");
            }
        }
        string STR;
        private void metroTile1_Click(object sender, EventArgs e)
        {
            flag = true;
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                STR = dialog.FileName;
            }
            else
            { flag = false; }
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Вы уверены что хотите выйти?", "Подтверждение выхода", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                Close();
            }
             }
        public static bool flag = false;
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
                    MessageBox.Show("Заполните все поля", "Корренктность ввода");
                else

                {                // отправитель - устанавливаем адрес и отображаемое в письме имя
                    MailAddress from = new MailAddress(metroTextBox1.Text, metroTextBox3.Text);
                    // кому отправляем
                    MailAddress to = new MailAddress(metroTextBox4.Text);
                    // создаем объект сообщения
                    MailMessage m = new MailMessage(from, to);
                    // тема письма
                    m.Subject = metroTextBox5.Text;
                    // текст письма
                    m.Body = richTextBox1.Text;
                    if (flag == true)
                    {
                        flag = false;  // !!! Важно 
                        m.Attachments.Add(new Attachment(STR = null));
                        // здесь свой код 
                    }

                    // письмо представляет код html
                    m.IsBodyHtml = true;
                    // адрес smtp-сервера и порт, с которого будем отправлять письмо
                    SmtpClient smtp = new SmtpClient("smtp.mail.ru", 587);
                    // логин и пароль
                    smtp.Credentials = new NetworkCredential(metroTextBox1.Text, metroTextBox2.Text);
                    smtp.EnableSsl = true;
                    smtp.Send(m);
                    MessageBox.Show("Сообщение отправлено на почту" + metroTextBox4.Text);
                }
            }
            catch
            {
                MessageBox.Show("Сообщение не отправлено, проверьте введенные данные");
            }
        }
    }
}
