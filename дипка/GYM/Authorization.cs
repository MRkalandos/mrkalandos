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

namespace GYM
{

    public partial class Authorization : Form
    {
        public Authorization()
        {
            InitializeComponent();
        }
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        private void Form1_Load(object sender, EventArgs e)
        {
            //     OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            //   con1.Open(); OleDbCommand sss1 = new OleDbCommand(@"select *  
            //                                                           from [зарплата_сотрудника] 
            //                                                         where зарплата=@st1 ", con1);
            /// sss1.Parameters.AddWithValue("st1", textBox1.Text);
            // sss1.ExecuteNonQuery();
            // if (sss1.ExecuteScalar() != null)
            //{
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    OleDbCommand Auth = new OleDbCommand("SELECT * FROM сотрудник WHERE должность = ? AND пароль = ?", con);                    
                    Auth.Parameters.AddWithValue("должность", comboBox1.Text);
                    Auth.Parameters.AddWithValue("пароль", textBox2.Text);
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
              
                    Auth1.Parameters.AddWithValue("должность", comboBox1.Text);
                    Auth1.Parameters.AddWithValue("пароль", textBox2.Text);
                    if (Auth1.ExecuteScalar() != null)
                    {
                        MessageBox.Show("Вход выполнен:Тренер");
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
               
 
            
           
            
            }
        }
    }
}
