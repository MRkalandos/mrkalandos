using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModTrener : MetroFramework.Forms.MetroForm
    {
        private const string Message = @"Неверный тип данных";
        private const string Title = @"Корректность ввода";
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Mod_trener.txt";

        public ModTrener()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_Trener_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox1.Text == "") ||
                    (textBox2.Text == "") ||
                    (textBox3.Text == "") ||
                    (maskedTextBox1.Text == "") ||
                    (metroTextBox4.Text == "") ||
                    (metroTextBox5.Text == "") ||
                    (metroTextBox1.Text == ""))
                {
                    MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    var connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                         Directory.GetParent(Directory.GetCurrentDirectory()).Parent
                                                             ?.FullName +
                                                         "/ISgym.mdb;Jet OLEDB:Database Password=316206");
                    connection.Open();
                    var queryFindClone = new OleDbCommand(@"select *  
                                                                      from [тренер] 
                                                                      where пароль=@pass and
                                                                    идтренер <> " + Convert.ToInt32(label1.Text) + "", connection);
                    queryFindClone.Parameters.AddWithValue("pass", metroTextBox5.Text);
                    queryFindClone.ExecuteNonQuery();
                    if (queryFindClone.ExecuteScalar() != null)
                    {
                        connection.Close();
                        MetroMessageBox.Show(this, "\nТакой пароль уже существует", TitleException,
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        this.DialogResult = DialogResult.OK;
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

        private void metroButton1_Click(object sender, EventArgs e)
        {
            var filedialog = new OpenFileDialog {Filter = @"JPG (*.jpg)|*.jpg|PNG (*.png)|*.png"};
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                var fileName = filedialog.SafeFileName;
                var sourcePath = filedialog.FileName;
                var targetPath = @"PhotoTrener";
                if (!Directory.Exists(targetPath))
                    Directory.CreateDirectory(targetPath);
                var destFile = Path.Combine(targetPath, fileName);
                try
                {
                    File.Copy(sourcePath, destFile, true);
                    metroTextBox6.Text = Path.GetFileName(fileName);
                }

                catch
                {
                    MetroMessageBox.Show(this, "\nПапка задействована, фото может быть не установлено", TitleException,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    metroTextBox6.Text = Path.GetFileName(fileName);
                }
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void MOD_Trener_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Mod_trener.chm"))
                {
                    Help.ShowHelp(null, "Help/Mod_trener.chm");
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

        private void MOD_Trener_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModTrener_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F6:
                    metroButton1.PerformClick();
                    break;
                case Keys.F5:
                    metroTile1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    metroTile2.PerformClick();
                    break;
            }
        }
    }
}
