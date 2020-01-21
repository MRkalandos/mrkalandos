using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class Money : MetroFramework.Forms.MetroForm
    {
        public Money()
        {
            InitializeComponent();
        }
        public DataTable dt;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
        public OleDbDataAdapter sda;

        public void upd()
        {
            try
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
                 catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Таблица не загрузилась");
            }
        }


        private void Money_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            upd();
        }

        private void Money_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string id;
            Money mon = new Money();
            MOD_Money ObjMoneyAdd = new MOD_Money();
            ObjMoneyAdd.textBox1.Text = "";
            ObjMoneyAdd.Text = "Добавить зарплату";
            ObjMoneyAdd.button1.Text = "Добавить";
            if (ObjMoneyAdd.ShowDialog() == DialogResult.OK)
                try
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
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this,ex.Message,"Запись не добавлена");
                }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con.Close();
            int rez = 0;
            MOD_Money ObjMoneyUpdate = new MOD_Money();

            ObjMoneyUpdate.Text = "Редактировать зарплату";
            ObjMoneyUpdate.button1.Text = "Редактировать";
            rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            ObjMoneyUpdate.textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            if (ObjMoneyUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                     con.Open();
                    OleDbCommand sss1 = new OleDbCommand(@"select *  
                                                          from [зарплата_сотрудника] 
                                                           where зарплата=@st1 ", con);
                    sss1.Parameters.AddWithValue("st1", ObjMoneyUpdate.textBox1.Text);
                    sss1.ExecuteNonQuery();
                    if (sss1.ExecuteScalar() != null)
                    {
                        con.Close();
                        MetroFramework.MetroMessageBox.Show(this, "\nТакая зарплата уже существует", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        con.Close();
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
                    MessageBox.Show(ex.Message,"Запись не изменена");
                }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Close();
                int rez = 0;
                if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
                {
                    con.Open();
                    rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                    OleDbCommand sss = new OleDbCommand(@"DELETE FROM зарплата_сотрудника 
                                                    WHERE идзарплата=" + rez + "", con);
                    sss.ExecuteNonQuery();
                    dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                    upd();
                    con.Close();
                }

            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись нельзя, данная запись используется в таблице 'Cотрудник'", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
