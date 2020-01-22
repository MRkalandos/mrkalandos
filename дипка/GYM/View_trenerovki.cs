using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class View_trenerovki : MetroFramework.Forms.MetroForm
    {
        public View_trenerovki()
        {
            InitializeComponent();
        }
        public int numbstrokaView = 0;
        public string idView;
        public DataTable dtView;
        public OleDbDataAdapter sdaview;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");



   
         public void upd()
        {
            try
            {
                sdaview = new OleDbDataAdapter(@"SELECT Вид_тренировки.* FROM Вид_тренировки;", con);
                dtView = new DataTable();
                sdaview.Fill(dtView);
                metroGrid1.DataSource = dtView;
                metroGrid1.Columns[0].Visible = false;
                metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                metroGrid1.Select();
                metroGrid1.AllowUserToAddRows = false;
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Таблица не загрузилась");
            }
        }

        private void View_trenerovki_Load(object sender, EventArgs e)
        {
            upd();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Close();
                int rez = 0;
                if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    con.Open();
                    rez = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
                    OleDbCommand sss = new OleDbCommand(@"DELETE FROM Вид_тренировки 
                                                    WHERE идвидтренировка=" + rez + "", con);
                    sss.ExecuteNonQuery();
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    upd();
                    con.Close();
                }

            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись нельзя, данная запись используется в таблице 'тренировка'", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string id;
            Money mon = new Money();
            MOD_View_trenerovka ObjTrenirovkaAdd = new MOD_View_trenerovka();
            ObjTrenirovkaAdd.textBox1.Text = "";
            ObjTrenirovkaAdd.Text = "Добавить тренировку";
            ObjTrenirovkaAdd.button1.Text = "Добавить";
            if (ObjTrenirovkaAdd.ShowDialog() == DialogResult.OK)
                try
                {

                    id = Convert.ToString(Convert.ToInt32(metroGrid1.Rows[metroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [Вид_тренировки]
                                                        ( название)
                                                        VALUES(@st1)", con);
                    sss.Parameters.AddWithValue("st1", ObjTrenirovkaAdd.textBox1.Text);
                    sss.ExecuteNonQuery();
                    con.Close();
                    upd();

                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Запись не добавлена");
                }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            con.Close();
            int rez = 0;
            MOD_View_trenerovka ObjMoneyUpdate = new MOD_View_trenerovka();

            ObjMoneyUpdate.Text = "Редактировать тренировку";
            ObjMoneyUpdate.button1.Text = "Редактировать";
            rez = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
            ObjMoneyUpdate.textBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[1].Value);
            if (ObjMoneyUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                        con.Close();
                        metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                        con.Open();
                        OleDbCommand sss = new OleDbCommand("update Вид_тренировки set название=@st1 where идвидтренировка=" + rez + "", con);
                        sss.Parameters.AddWithValue("st1", ObjMoneyUpdate.textBox1.Text);
                        sss.ExecuteNonQuery();
                        con.Close();
                        upd();
                    }
                
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Запись не изменена");
                }
        }
    }
}
