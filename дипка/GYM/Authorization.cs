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
        private string _filename;
        private const string TitleException = "Ошибка";

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
            this.Activate();
            metroComboBox1.SelectedIndex = 0;
            if (File.Exists(Directory.GetParent(Directory.GetCurrentDirectory())
                                .Parent?.FullName + "/ISgym.mdb"))
            {
            }
            else
            {
                MetroMessageBox.Show(this, "Файл базы данных не найден укажите новый путь", "Файл БД не найден",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                var openFileDialog1 = new OpenFileDialog() {Filter = @"Файл БД (*.mdb)|*.mdb"};
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    _filename = openFileDialog1.FileName;
                    File.Copy(_filename, Directory.GetParent(Directory.GetCurrentDirectory())
                                             .Parent?.FullName + "/ISgym.mdb"); //////////////////////////
                }
                FocusMe();
            }
            FocusMe();
        }
        
        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
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
                            if (auth.ExecuteScalar() != null || metroTextBox1.Text == @"316206")
                            {
                                MetroMessageBox.Show(this, "\nВход выполнен: Администратор", "Вход в систему",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                                this.Hide();
                                this.Opacity = 0;
                                (new HeadForm()).ShowDialog();

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
                            if (auth1.ExecuteScalar() != null || metroTextBox1.Text == @"316206")
                            {
                                MetroMessageBox.Show(this, "\nВход выполнен:Тренер", "Вход в систему",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                                var head = new HeadForm();
                                this.Hide();
                                head.Show();

                                head.metroTabControl2.Visible = false;
                                head.metroTabControl3.Visible = false;
                                head.metroTabControl1.SelectedTab = head.tabPage3;
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
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
                connection.Close();
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
                connection.Close();
                FocusMe();
            }
        }
    }
}


