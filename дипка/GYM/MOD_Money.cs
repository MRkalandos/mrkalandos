using System;
using System.Data.OleDb;
using MetroFramework.Forms;
using System.Windows.Forms;
using System.IO;
using MetroFramework;

namespace GYM
{
    public partial class ModMoney : MetroForm
    {
        public static OleDbConnection connection = new OleDbConnection(ConnectionDb.conString());
        private const string TitleException = "Ошибка";
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
            else
            {
                FocusMe();
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
                    MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    connection.Open();
                    var sss1 = new OleDbCommand(@"select зарплата  
                                                 from [зарплата_сотрудника] 
                                                     where зарплата=@st1 
                                                       and идзарплата <> " + Convert.ToInt32(metroLabel1.Text) + "",
                        connection);
                    sss1.Parameters.AddWithValue("st1", textBox1.Text);
                    sss1.ExecuteNonQuery();
                    if (sss1.ExecuteScalar() != null)
                    {
                        MetroMessageBox.Show(this, "\nТакая зарплата уже существует", "Корректность",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        this.DialogResult = DialogResult.OK;
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
                    connection.Close();
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

        private void ModMoney_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModMoney_FormClosing(object sender, FormClosingEventArgs e)
        {
  
        }
    }
}
