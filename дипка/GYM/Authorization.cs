using System;
using System.Data.OleDb;
using System.Windows.Forms;
using MetroFramework.Forms;

namespace GYM
{

    public partial class Authorization : MetroForm
    {
        public Authorization()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            metroComboBox1.SelectedIndex = 0;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((metroTextBox1.Text == ""))

            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗаполните поле пароль", "Поле пустое", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
            else
            {
                con.Open();
                OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
                switch (metroComboBox1.SelectedIndex)
                {
                    case 0:
                        OleDbCommand Auth = new OleDbCommand("SELECT * FROM сотрудник WHERE должность = ? AND пароль = ?", con);
                        Auth.Parameters.AddWithValue("должность", metroComboBox1.Text);
                        Auth.Parameters.AddWithValue("пароль", metroTextBox1.Text);
                        if (Auth.ExecuteScalar() != null)
                        {
                            MetroFramework.MetroMessageBox.Show(this, "\nВход выполнен: Администратор", "Вход в систему", MessageBoxButtons.OK, MessageBoxIcon.Information);
                           
                            HeadForm Head = new HeadForm();
                            Head.Show();
                            con.Close();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "\nНе верный Пароль", "Ошибка входа", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            con.Close();
                        }
                        break;
                    case 1:
                        OleDbCommand Auth1 = new OleDbCommand("SELECT * FROM тренер WHERE должность = ? AND пароль = ?", con);

                        Auth1.Parameters.AddWithValue("должность", metroComboBox1.Text);
                        Auth1.Parameters.AddWithValue("пароль", metroTextBox1.Text);
                        if (Auth1.ExecuteScalar() != null)
                        {
                            MetroFramework.MetroMessageBox.Show(this, "\nВход выполнен:Тренер", "Вход в систему", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            HeadForm Head = new HeadForm();
                            Head.Show();
                            //Head.tabControl2.Visible = false;
                            //Head.tabControl3.Visible = false;
                            //Head.tabControl1.SelectedTab = Head.tabPage5;
                            //Head.pictureBox2.Visible = true;
                            //Head.pictureBox3.Visible = true;
                            con.Close();
                        }
                        else
                        {
                            MetroFramework.MetroMessageBox.Show(this, "\nНе верный Пароль", "Ошибка входа", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            con.Close();
                        }
                        break;
                }
            }
        }

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {

                if (e.KeyChar != (char)Keys.Back )
                    if (e.KeyChar != (char)Keys.Enter)
                    {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНе верный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Authorization_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                metroTile1.PerformClick();// имитируем нажатие button1
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
         Help.ShowHelp(null, "Help/Сотрудники.chm");

        }
    }
}

