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
    public partial class VISITS : MetroFramework.Forms.MetroForm
    {
        public VISITS()
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
                sda = new OleDbDataAdapter(@"SELECT  Учет_посещений.Идпосещений, Спортсмен.Фамилия, Абонемент.Название, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Учет_посещений.Дата,Учет_посещений.идпродажа
FROM Абонемент INNER JOIN ((Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) INNER JOIN Учет_посещений ON Продажа_абонемента.Идпродажа = Учет_посещений.Идпродажа) ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент;", con);
                dt = new DataTable();
                sda.Fill(dt);
                metroGrid1.DataSource = dt;
                metroGrid1.Columns[0].Visible = false;
                metroGrid1.Columns[6].Visible = false;
                metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                metroGrid1.Select();
                metroGrid1.AllowUserToAddRows = false;
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Таблица не загрузилась");
            }
        }


        private void VISITS_Load(object sender, EventArgs e)
        {
            upd();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            try
            {
                con.Close();
                int rez = 0;
                if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?", MessageBoxButtons.YesNo, MessageBoxIcon.Information))
                {
                    con.Open();
                    rez = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
                    OleDbCommand sss = new OleDbCommand(@"DELETE FROM учет_посещений 
                                                    WHERE идпосещений=" + rez + "", con);
                    sss.ExecuteNonQuery();
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    upd();
                    con.Close();
                }

            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись нельзя, данная запись используется в таблице 'Cотрудник'", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public string idvisit;
        private void metroButton1_Click(object sender, EventArgs e)
        {
            MOD_VISITS ObjVISITeAdd = new MOD_VISITS();
           
            ObjVISITeAdd.Text = "Добавить посещение";
            ObjVISITeAdd.metroTile1.Text = "Добавить";
            con.Open();
            OleDbCommand cmd = new OleDbCommand(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия
FROM Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;", con);
            ObjVISITeAdd.metroComboBox1.DisplayMember = "Спортсмен.Фамилия";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjVISITeAdd.metroComboBox1.DataSource = list.ToList();
            ObjVISITeAdd.metroComboBox1.DisplayMember = "Value";
            ObjVISITeAdd.metroComboBox1.ValueMember = "Key";
            con.Close();
            if (ObjVISITeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    idvisit = Convert.ToString(Convert.ToInt32(metroGrid1.Rows[metroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [учет_посещений]
                                                        ( Дата,Идпродажа)
                                                        VALUES(@st1,@st2)", con);

                    sss.Parameters.AddWithValue("st1", Convert.ToDateTime(ObjVISITeAdd.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjVISITeAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.ExecuteNonQuery();
                    con.Close();
                    upd();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }
        public int numbstr;
        private void metroButton2_Click(object sender, EventArgs e)
        {
            MOD_VISITS ObjVisiteUpdate = new MOD_VISITS();
            con.Close();
            ObjVisiteUpdate.Text = "Редактировать посещение";
            ObjVisiteUpdate.metroTile1.Text = "Редактировать";
            numbstr = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
           
            ObjVisiteUpdate.metroDateTime1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[5].Value);
            con.Open();
            OleDbCommand cmd = new OleDbCommand(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия
FROM Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;", con);
            ObjVisiteUpdate.metroComboBox1.DisplayMember = "Спортсмен.Фамилия";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjVisiteUpdate.metroComboBox1.DataSource = list.ToList();
            ObjVisiteUpdate.metroComboBox1.DisplayMember = "Value";
            ObjVisiteUpdate.metroComboBox1.ValueMember = "Key";
            con.Close();
            ObjVisiteUpdate.metroComboBox1.SelectedValue = metroGrid1.CurrentRow.Cells[6].Value;
            if (ObjVisiteUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    con.Close();
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update Учет_посещений set дата=@st1, идпродажа=@st2 where идпосещений=" + numbstr + "", con);
                  sss.Parameters.AddWithValue("st1", Convert.ToDateTime(ObjVisiteUpdate.metroDateTime1.Text));
              
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjVisiteUpdate.metroComboBox1.SelectedValue));
                    sss.ExecuteNonQuery();
                    con.Close();
                    upd();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }
    }
}
