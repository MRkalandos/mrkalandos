using System;
using System.Data.OleDb;
using MetroFramework.Forms;
using System.Windows.Forms;
using System.IO;

namespace GYM
{
    public partial class ModMoney : MetroForm
    {
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Mod_money.txt";

        public ModMoney()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this,
                    "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
                connection.Open();
                var sss1 = new OleDbCommand(@"select зарплата  
                                                 from [зарплата_сотрудника] 
                                                     where зарплата=@st1 
                                                       and идзарплата <> " + Convert.ToInt32(metroLabel1.Text) + "", connection);
                sss1.Parameters.AddWithValue("st1", textBox1.Text);
                sss1.ExecuteNonQuery();
                if (sss1.ExecuteScalar() != null)
                {
                    connection.Close();
                    MetroFramework.MetroMessageBox.Show(this, "\nТакая зарплата уже существует", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void AddMoney_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Mod_money.chm"))
                {
                    Help.ShowHelp(null, "Help/Mod_money.chm");
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Файл не найден", TitleException,MessageBoxButtons.OK,MessageBoxIcon.Error);
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

        private void ModMoney_Activated(object sender, EventArgs e)
        {
            Focus();
        }

        private void ModMoney_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModMoney_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5:
                    button1.PerformClick();
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
