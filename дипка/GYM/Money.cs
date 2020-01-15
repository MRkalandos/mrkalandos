using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class Money : Form
    {
        public Money()
        {
            InitializeComponent();
        }
        public DataTable dt;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        public OleDbDataAdapter sda;

        public void upd()
        {
            sda = new OleDbDataAdapter(@"SELECT * from зарплата_сотрудника;", con);
            dt = new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
            dataGridView1.Select();
            dataGridView1.AllowUserToAddRows = false;
        }

        private void Money_Load(object sender, EventArgs e)
        {
            upd();
        }

        private void Money_FormClosed(object sender, FormClosedEventArgs e)
        {
            AddMoney money = new AddMoney();

         //  this.зарплата_сотрудникаTableAdapter1.Update(this.dS_Money1.Зарплата_сотрудника);
           // this.зарплата_сотрудникаTableAdapter1.Fill(this.dS_Money1.Зарплата_сотрудника);
            // HeadForm main1 = new HeadForm();
            //main1.Opacity = 1; // возвращаем нормальный вид главной форме
        }

        private void button1_Click(object sender, EventArgs e)
        {
       string id;
        Money mon = new Money();
            AddMoney ObjMoneyAdd = new AddMoney();
            ObjMoneyAdd.textBox1.Text = "";
            ObjMoneyAdd.Text = "Добавить зарплату";
            ObjMoneyAdd.button1.Text = "Добавить";
            if (ObjMoneyAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
                   con1.Open();  OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                                      from [зарплата_сотрудника] 
                                                                      where зарплата=@st1 ", con1);
                    sss1.Parameters.AddWithValue("st1", ObjMoneyAdd.textBox1.Text);
                    sss1.ExecuteNonQuery(); 
                    if (sss1.ExecuteScalar() != null)
                    {
                        con1.Close();
                        MessageBox.Show("Запись существует"); 
                    }
                    else
                    {
                        id = Convert.ToString(Convert.ToInt32(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value) + 1);
                        con.Open();
                        OleDbCommand sss = new OleDbCommand(@"INSERT INTO [Зарплата_сотрудника]
                                                        ( Зарплата)
                                                        VALUES(@st1)", con);
                        sss.Parameters.AddWithValue("st1", ObjMoneyAdd.textBox1.Text);
                        sss.ExecuteNonQuery();
                        con.Close();
                        upd();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int rez = 0;
         //   string id;
            AddMoney ObjMoneyUpdate = new AddMoney();
         
            ObjMoneyUpdate.Text = "Редактировать зарплату";
            ObjMoneyUpdate.button1.Text = "Редактировать";
            rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            ObjMoneyUpdate.textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            if (ObjMoneyUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    OleDbConnection con1 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
                    con1.Open(); OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                                      from [зарплата_сотрудника] 
                                                                      where зарплата=@st1 ", con1);
                    sss1.Parameters.AddWithValue("st1", ObjMoneyUpdate.textBox1.Text);
                    sss1.ExecuteNonQuery();
                    if (sss1.ExecuteScalar() != null)
                    {
                        con1.Close();
                        MessageBox.Show("Запись существует");
                    }
                    else
                    {
                        dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                        con.Open();
                        OleDbCommand sss = new OleDbCommand("update зарплата_сотрудника set зарплата=@st1 where идзарплата=" + rez + "", con);
                        sss.Parameters.AddWithValue("st1", ObjMoneyUpdate.textBox1.Text);
                        sss.ExecuteNonQuery();
                        con.Close();
                        upd();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int rez=0;
            if (DialogResult.Yes == MessageBox.Show("Вы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                con.Open();
                rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM зарплата_сотрудника 
                                                    WHERE идзарплата=" + rez + "", con);
                sss.ExecuteNonQuery();
                upd();
                con.Close();
            }
        }
    }
}
