using System;
using System.Data.OleDb;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.IO;
using MetroFramework;

namespace GYM
{
    public partial class Authorization : MetroForm
    {
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Authorization.txt";
        public OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                                Directory.GetParent(Directory.GetCurrentDirectory())
                                                                    .Parent?.FullName +
                                                                "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public Authorization()
        {
            InitializeComponent();
            KeyPreview = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FormBorderStyle = FormBorderStyle.None;
            metroComboBox1.SelectedIndex = 0;
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try { 
                if (metroTextBox1.Text == "")

                {
                    MetroMessageBox.Show(this, "\nЗаполните поле пароль", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                else
                {
                    connection.Open();
                    switch (metroComboBox1.SelectedIndex)
                    {
                        case 0:
                            var auth = new OleDbCommand("SELECT * FROM сотрудник WHERE должность = ? AND пароль = ?",
                                connection);
                            auth.Parameters.AddWithValue("должность", metroComboBox1.Text);
                            auth.Parameters.AddWithValue("пароль", metroTextBox1.Text);
                            if (auth.ExecuteScalar() != null)
                            {
                                MetroMessageBox.Show(this, "\nВход выполнен: Администратор", "Вход в систему",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                                var head = new HeadForm();
                                head.Show();
                                connection.Close();
                            }
                            else
                            {
                                MetroMessageBox.Show(this, "\nНе верный Пароль", TitleException, MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);

                                connection.Close();
                            }

                            break;
                        case 1:
                            var auth1 = new OleDbCommand("SELECT * FROM тренер WHERE должность = ? AND пароль = ?",
                                connection);
                            auth1.Parameters.AddWithValue("должность", metroComboBox1.Text);
                            auth1.Parameters.AddWithValue("пароль", metroTextBox1.Text);
                            if (auth1.ExecuteScalar() != null)
                            {
                                MetroMessageBox.Show(this, "\nВход выполнен:Тренер", "Вход в систему", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                                var head = new HeadForm();
                                head.Show();
                                head.metroTabControl2.Visible = false;
                                head.metroTabControl3.Visible = false;
                                head.metroTabControl1.SelectedTab = head.tabPage3;
                                //Head.tabControl2.Visible = false;
                                //Head.tabControl3.Visible = false;
                                //Head.tabControl1.SelectedTab = Head.tabPage5;
                                //Head.pictureBox2.Visible = true;
                                //Head.pictureBox3.Visible = true;
                                connection.Close();
                            }
                            else
                            {
                                MetroMessageBox.Show(this, "\nНе верный Пароль", TitleException, MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                                connection.Close();
                            }

                            break;
                    }
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            var blockCifr = e.KeyChar;
            if (blockCifr >= '0' && blockCifr <= '9') return;
            if (e.KeyChar == (char) Keys.Back) return;
            if (e.KeyChar == (char) Keys.Enter) return;
            e.Handled = true;
            MetroMessageBox.Show(this, "\nНе верный тип данных", TitleException,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Authorization_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    metroTile1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    Environment.Exit(0);
                    break;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Authorization.chm"))
                {
                    Help.ShowHelp(null, "Help/Authorization.chm");
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
    }
}

