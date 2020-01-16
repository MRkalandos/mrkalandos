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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        OleDbDataAdapter ada = new OleDbDataAdapter("Select * From сотрудник where Должность = '" + comboBox1.Text + "'and Пароль = '" +Convert.ToInt32( textBox2.Text) + "'", con);
            DataTable dt = new DataTable();
            ada.Fill(dt);
            if (dt.Rows[0][0].ToString() == "1")
            {
                MessageBox.Show("");
                //Hide();
            //    FrmGlav fGl = new FrmGlav();
               // fGl.Show();
            }
            else
            {
                MessageBox.Show("Неправильный логин или пароль!!!");
            }

        }
    }
}
