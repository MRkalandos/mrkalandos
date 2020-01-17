using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;


namespace GYM
{

    public partial class Authorization : MetroForm
    {
        public Authorization()
        {
            InitializeComponent();
        }
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            
           // this.Style = MetroFramework.MetroColorStyle.Red;
            // this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            metroComboBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            //    MessageBox.Show(this, "Your message here.", "Title Here", MessageBoxButtons.OKCancel, MessageBoxIcon.Hand);
        }

        private void metroPanel1_Paint(object sender, PaintEventArgs e)
        {

        }




        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((metroTextBox1.Text == ""))

            {
                MessageBox.Show("Заполните поле пароль,Поле пустое");
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
                            MessageBox.Show("Вход выполнен:Администратор");
                            HeadForm Head = new HeadForm();
                            Head.Show();
                            con.Close();
                        }
                        else
                        {
                            MessageBox.Show("Не верный пользователь или пароль");
                            con.Close();
                        }
                        break;
                    case 1:
                        OleDbCommand Auth1 = new OleDbCommand("SELECT * FROM тренер WHERE должность = ? AND пароль = ?", con);

                        Auth1.Parameters.AddWithValue("должность", metroComboBox1.Text);
                        Auth1.Parameters.AddWithValue("пароль", metroTextBox1.Text);
                        if (Auth1.ExecuteScalar() != null)
                        {
                            MessageBox.Show("Вход выполнен:Тренер");
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
                            MessageBox.Show("Не верный пользователь или пароль");
                            con.Close();
                        }
                        break;
                }


            }
        }
    }
}

