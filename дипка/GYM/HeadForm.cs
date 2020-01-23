using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GYM
{
    public partial class HeadForm : MetroFramework.Forms.MetroForm
    {
        public HeadForm()
        {
            InitializeComponent();
           
            metroTile4.PerformClick();
        }
        public int numbstrokasportsmen = 0;
        public int numbstrokTrenerovka = 0;
        public int numbstrokemployee = 0;
        public int numbstroktrener = 0;
        public int numbstrSALE = 0;
        public int numbAbonement = 0;
        public string idEmployee;
        public string idSALE;
        public string idsportsmen;
        public string idtrener;
        public string idabonement;
        public string idtrenerovka;
        public DataTable dtEmployee;
        public DataTable dtSportsmen;
        public DataTable dtSALE;
        public DataTable dtAbonement;
        public DataTable dtTrener;
        public DataTable dtTrenerovka;
        string filename;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
       public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
        public OleDbDataAdapter sdaEmployee;
        public OleDbDataAdapter sdaSALE;
        public OleDbDataAdapter sdasportsmen;
        public OleDbDataAdapter sdaabonement;
        public OleDbDataAdapter sdatrener;
        public OleDbDataAdapter sdatrenerovka;
        public void updSportsmen()
        {
            try
            {
                sdasportsmen = new OleDbDataAdapter(@"SELECT Спортсмен.*
                                                    FROM Спортсмен;", con);
                dtSportsmen = new DataTable();
                sdasportsmen.Fill(dtSportsmen);
                SPORTMmetroGrid2.DataSource = dtSportsmen;
                SPORTMmetroGrid2.Columns[0].Visible = false;
                SPORTMmetroGrid2.Select();
                SPORTMmetroGrid2.AllowUserToAddRows = false;
                SPORTMmetroTabControl12.Enabled = true;
                metroTabControl2.Enabled = true;
                SPORTMmetroTabControl10.Enabled = true;
                SPORTMmetroTabControl9.Enabled = true;
                metroTile33.Enabled = true;
                metroTile32.Enabled = true;
                metroTile34.Enabled = true;
                SPORTMmetroTabControl9.Enabled = true;
                metroTile30.Enabled = true;
                SPORTMtextBox1.Text = "";
                SPORTMtextBox6.Text = "";
                SPORTMtextBox7.Text = "";
                SPORTMtextBox9.Text = "";
                SPORTMtextBox10.Text = "";
                metroTile29.Enabled = true;
                metroTile31.Enabled = true;
                SPORTMmetroTabControl11.Enabled = true;
                SPORTMmetroTabControl14.Enabled = true;
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void updTrenerovka()
        {
            try
            {
                sdatrenerovka = new OleDbDataAdapter(@"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер;
", con);
                dtTrenerovka = new DataTable();
                sdatrenerovka.Fill(dtTrenerovka);
                TRENINGmetroGrid1.DataSource = dtTrenerovka;
                TRENINGmetroGrid1.Columns[0].Visible = false;
                TRENINGmetroGrid1.Columns[4].Visible = false;
                TRENINGmetroGrid1.Columns[5].Visible = false;
                TRENINGmetroGrid1.Select();
                TRENINGmetroGrid1.AllowUserToAddRows = false;
                TRENINGmetroTabControl6.Enabled = true;
                metroTile18.Enabled = true;
                TRENINGmetroTabControl4.Enabled = true;
                TRENINGmetroTabControl3.Enabled = true;        
                TRENINGmetroTabControl5.Enabled = true;
                TRENINGmetroTabControl7.Enabled = true;
                TRENINGmetroTabControl8.Enabled = true;
                TRENINGtextBox6.Text = "";
                TRENINGtextBox5.Text = "";
                TRENINGtextBox4.Text = "";
                TRENINGtextBox3.Text = "";
                TRENINGmetroTabControl5.Enabled = true;
                metroTile15.Enabled = true;
                metroTile14.Enabled = true;
                metroTile17.Enabled = true;

            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void updSale()
        {
            try
            {
                sdaSALE = new OleDbDataAdapter(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN (Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник;", con);
                dtSALE = new DataTable();
                sdaSALE.Fill(dtSALE);
                SALEmetroGrid1.DataSource = dtSALE;
                SALEmetroGrid1.Columns[0].Visible = false;
                SALEmetroGrid1.Columns[7].Visible = false;
                SALEmetroGrid1.Columns[8].Visible = false;
                SALEmetroGrid1.Columns[9].Visible = false;
                SALEmetroTabControl4.Enabled = true;
                metroTile44.Enabled = true;
                SALEmetroTabControl6.Enabled = true;
                SALEmetroTabControl9.Enabled = true;
                SALEmetroTabControl7.Enabled = true;
                SALEmetroTabControl5.Enabled = true;
                SALEmetroTabControl7.Enabled = true;
                metroTile42.Enabled = true;
                metroTile43.Enabled = true;
                SALEmetroGrid1.Select();
                SALEmetroGrid1.AllowUserToAddRows = false;
                metroTile40.Enabled = true;
                metroTile41.Enabled = true;
                SALEtextBox6.Text = "";
                SALEtextBox4.Text = "";
                SALEtextBox5.Text = "";
                SALEmetroTextBox2.Text = "";
                SALEmetroTextBox3.Text = "";
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void updAbonement()
        {
            try
            {
                sdaabonement = new OleDbDataAdapter(@"SELECT Абонемент.Идабонемент, Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка;", con);
                dtAbonement = new DataTable();
                sdaabonement.Fill(dtAbonement);
                ABONmetroGrid1.DataSource = dtAbonement;
                ABONmetroGrid1.Columns[0].Visible = false;
                ABONmetroGrid1.Columns[5].Visible = false;
                ABONmetroTabControl7.Enabled = true;
                ABONmetroTabControl5.Enabled = true;
                ABONmetroTabControl4.Enabled = true;
                metroTile27.Enabled = true;
                metroTile35.Enabled = true;
                ABONmetroTabControl7.Enabled = true;
                ABONmetroTabControl6.Enabled = true;
                ABONmetroTabControl4.Enabled = true;
                metroTile19.Enabled = true;
                metroTile5.Enabled = true;
              
                metroTile26.Enabled = true;


                ABONmetroGrid1.Select();
                ABONmetroGrid1.AllowUserToAddRows = false;
                }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        public void updTrener()
        {
            try
            {
                sdatrener = new OleDbDataAdapter(@"SELECT Тренер.*
                                                    FROM Тренер;", con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                TRENmetroGrid1.DataSource = dtTrener;
                TRENmetroGrid1.Columns[0].Visible = false;
                TRENmetroGrid1.Columns[8].Visible = false;
                TRENtextBox8.Text = "";
                TRENtextBox5.Text = "";
                TRENtextBox4.Text = "";
                TRENtextBox2.Text = "";
                TRENtextBox3.Text = "";
                TRENmetroTextBox2.Text = "";
                TRENmetroTextBox3.Text = "";
                TRENmetroGrid1.Select();
                TRENmetroGrid1.AllowUserToAddRows = false;
                metroTile16.Visible = true;
                TRENpictureBox2.Visible = true;
                metroTile11.Enabled = true;
                metroTile12.Enabled = true;
                metroTile10.Enabled = true;
                TRENmetroTabControl4.Enabled = true;
                TRENmetroTabControl3.Enabled = true;
                TRENmetroTabControl6.Enabled = true;
                metroTile7.Enabled = true;
                metroTile8.Enabled = true;
                metroTile9.Enabled = true;
                TRENmetroTabControl5.Enabled = true;
                TRENmetroTabControl8.Enabled = true;
                
                TRENmetroTabControl5.Enabled = true;


            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void updEmployee()
        {
            try
            {
                sdaEmployee = new OleDbDataAdapter(@"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата, сотрудник.идзарплата  
                                        FROM Зарплата_сотрудника 
                                        INNER JOIN Сотрудник 
                                        ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", con);
                dtEmployee = new DataTable();
                sdaEmployee.Fill(dtEmployee);
                EMPLmetroGrid1.DataSource = dtEmployee;
                EMPLmetroTile3.Visible = true;
                EMPLtextBox2.Text = "";
                EMPLtextBox3.Text = "";
                EMPLtextBox4.Text = "";
                EMPLtextBox5.Text = "";
                EMPLtextBox8.Text = "";
                EMPLmetroTextBox6.Text = "";
                EMPLmetroTextBox7.Text = "";
                EMPLmetroTabControl3.Enabled = true;
                EMPLmetroTabControl4.Enabled = true;
                EMPLmetroTabControl5.Enabled = true;
                EMPLmetroTabControl6.Enabled = true;
                EMPLmetroTabControl5.Enabled = true;
                EMPLmetroTabControl7.Enabled = true;
                EMPLmetroTabControl8.Enabled = true;
                EMPLmetroTile10.Enabled = true;
                EMPLmetroTile16.Enabled = true;
                EMPLmetroTabControl8.Enabled = true;
                EMPLmetroTabControl7.Enabled = true;
                EMPLmetroTabControl8.Enabled = true;
                EMPLmetroTile11.Enabled = true;
                EMPLmetroTile16.Enabled = true;
                EMPLmetroTile14.Enabled = true;
                EMPLmetroTile15.Enabled = true;
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroGrid1.Columns[0].Visible = false;
                EMPLmetroGrid1.Columns[10].Visible = false;
                EMPLmetroGrid1.Select();
                EMPLmetroGrid1.AllowUserToAddRows = false;
                /////////////////////////////sportsmen////////////////////////////////////////
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе удалось обновить таблицу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void updfilename()
        {
            OleDbConnection conEmployee = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdaEmployee = new OleDbDataAdapter(@"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата
                                       FROM Зарплата_сотрудника 
                                       INNER JOIN Сотрудник   
                                       ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", conEmployee);
            dtEmployee = new DataTable();
            sdaEmployee.Fill(dtEmployee);
            EMPLmetroGrid1.DataSource = dtEmployee;

            ////////////////////////////////////////////sportsmen//////////////////////////////////////////
            OleDbConnection conSportsmen = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdasportsmen = new OleDbDataAdapter(@"SELECT Спортсмен.*
                                                 FROM Спортсмен;", conSportsmen);
            dtSportsmen = new DataTable();
            sdasportsmen.Fill(dtSportsmen);
            SPORTMmetroGrid2.DataSource = dtSportsmen;
            /////////////////////////////////////////////abonement//////////////////////
            OleDbConnection conAbonement = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdaabonement = new OleDbDataAdapter(@"SELECT Абонемент.Идабонемент, Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
FROM Тренировка INNER JOIN Абонемент ON Тренировка.Идтренировка = Абонемент.Идтренеровка;
", conAbonement);
            dtAbonement = new DataTable();
            sdaabonement.Fill(dtAbonement);
            ABONmetroGrid1.DataSource = dtAbonement;

            ////////////////////////////////////////////////////////////sale///////////////////
            OleDbConnection conSale = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdaabonement = new OleDbDataAdapter(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN (Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник;", conSale);
            dtSALE = new DataTable();
            sdaSALE.Fill(dtSALE);
            SALEmetroGrid1.DataSource = dtSALE;

            ////////////////////////////////////trenerovka
            OleDbConnection contrenerovka = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdatrenerovka = new OleDbDataAdapter(@"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер;
", contrenerovka);
            dtTrenerovka = new DataTable();
            sdatrenerovka.Fill(dtTrenerovka);
            TRENINGmetroGrid1.DataSource = dtTrenerovka;

            ///////////////////////////////////////////////trener///////////////////////////////////
            OleDbConnection conTrener = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdatrener = new OleDbDataAdapter(@"SELECT тренер.*
                                                 FROM тренер;", conTrener);
            dtTrener = new DataTable();
            sdatrener.Fill(dtTrener);
            TRENmetroGrid1.DataSource = dtTrener;

            TRENmetroGrid1.Columns[0].Visible = false;
            SPORTMmetroGrid2.Columns[0].Visible = false;
            EMPLmetroGrid1.Columns[0].Visible = false;
            TRENINGmetroGrid1.Columns[0].Visible = false;
        }

        private void HeadeForm_Load(object sender, EventArgs e)
        {

            SPORTMmetroComboBox1.SelectedIndex = 0;
            
            try
            {
                updTrener();
                updEmployee();
                updSportsmen();
                updTrenerovka();
                updAbonement();
                updSale();
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "\nФайл базы данных не найден укажите новый путь", "Файл БД не найден", MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenFileDialog openFileDialog1 = new OpenFileDialog() { Filter = "Файл БД (*.mdb)|*.mdb" };
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (StreamWriter sw = File.CreateText(Application.StartupPath + @"\" + "text.txt"))
                    {
                        filename = openFileDialog1.FileName;
                        sw.WriteLine(filename);
                        File.Copy(filename, AppDomain.CurrentDomain.BaseDirectory + Path.GetFileName(filename));
                        updfilename();
                    }
                }
                else
                {
                    string path = Application.StartupPath + @"\" + "text.txt";
                    using (StreamReader sr = new StreamReader(path))
                    {
                        filename = sr.ReadLine();
                    }
                    updfilename();
                    MetroFramework.MetroMessageBox.Show(this, "\nВы не выбрали файл БД, поэтому выбранно последнее выбранное место БД", "Не указана БД", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void metroTile18_Click(object sender, EventArgs e)
        {
            
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECt сотрудник.идсотрудник, Сотрудник.Фамилия,Сотрудник.Имя,Сотрудник.отчество,Сотрудник.Дата,Сотрудник.Должность,Сотрудник.Телефон, '(' & Зарплата_сотрудника.Зарплата & ')' AS Зарплата ,сотрудник.фото
                                                                         FROM Зарплата_сотрудника 
                                                                         INNER JOIN Сотрудник 
                                                                         ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                                                                         UNION 
                                                                         SELECT 'Средняя:','', '','','','','', Avg(Зарплата_сотрудника.Зарплата) AS Средняя,'' 
                                                                         FROM Зарплата_сотрудника INNER JOIN Сотрудник 
                                                                         ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата "), conn);
            adapter.Fill(ds);
            EMPLmetroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            EMPLmetroGrid1.Columns[0].Visible = false;
            EMPLmetroGrid1.Columns[8].Visible = false;
            EMPLmetroTabControl3.Enabled = false;
            if (EMPLmetroGrid1.RowCount == 0)
            {

                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updEmployee();
            }
        }

        private void metroTile4_Click(object sender, EventArgs e)
        {
            MOD_Employee ObjEmployeeAdd = new MOD_Employee();
            ObjEmployeeAdd.textBox1.Text = "";
            ObjEmployeeAdd.textBox2.Text = "";
            ObjEmployeeAdd.textBox3.Text = "";
            ObjEmployeeAdd.maskedTextBox1.Text = "";
            ObjEmployeeAdd.metroDateTime1.Text = null;
           // ObjEmployeeAdd.metroTextBox6.Text = "";
            ObjEmployeeAdd.metroTextBox5.Text = "";
            ObjEmployeeAdd.Text = "Добавить сотрудника";
            ObjEmployeeAdd.metroTile1.Text = "Добавить";
            con.Open();
            OleDbCommand cmd = new OleDbCommand("select идзарплата,зарплата from зарплата_сотрудника", con);
            ObjEmployeeAdd.metroComboBox1.DisplayMember = "Зарплата";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, int> list = new Dictionary<int, int>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (int)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjEmployeeAdd.metroComboBox1.DataSource = list.ToList();
            ObjEmployeeAdd.metroComboBox1.DisplayMember = "Value";
            ObjEmployeeAdd.metroComboBox1.ValueMember = "Key";
            con.Close();
            if (ObjEmployeeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    EMPLmetroGrid1.Sort(EMPLmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    idEmployee = Convert.ToString(Convert.ToInt32(EMPLmetroGrid1.Rows[EMPLmetroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [Сотрудник]
                                                        ( Фамилия, Имя, Отчество, Должность,Телефон, Дата,Пароль,Фото,Идзарплата)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5,@st6,@st7,@st8,@st9)", con);
                    sss.Parameters.AddWithValue("st1", ObjEmployeeAdd.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjEmployeeAdd.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjEmployeeAdd.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjEmployeeAdd.metroTextBox4.Text);
                    sss.Parameters.AddWithValue("st5", ObjEmployeeAdd.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjEmployeeAdd.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st7", Convert.ToInt32(ObjEmployeeAdd.metroTextBox5.Text));
                    sss.Parameters.AddWithValue("st8", ObjEmployeeAdd.metroTextBox6.Text);
                    sss.Parameters.AddWithValue("st9", Convert.ToInt32(ObjEmployeeAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updEmployee();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            MOD_Employee ObjEmployeeUpdate = new MOD_Employee();
            con.Close();
            ObjEmployeeUpdate.Text = "Редактировать сотрудника";
            ObjEmployeeUpdate.metroTile1.Text = "Редактировать";
            numbstrokemployee = Convert.ToInt32(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
            ObjEmployeeUpdate.textBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[1].Value);
            ObjEmployeeUpdate.textBox2.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[2].Value);
            ObjEmployeeUpdate.textBox3.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[3].Value);
            ObjEmployeeUpdate.metroTextBox4.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[4].Value);
            ObjEmployeeUpdate.maskedTextBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[5].Value);
            ObjEmployeeUpdate.metroDateTime1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[6].Value);
            ObjEmployeeUpdate.metroTextBox5.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[7].Value);
            ObjEmployeeUpdate.metroTextBox6.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[8].Value);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("select идзарплата,зарплата from зарплата_сотрудника", con);
            ObjEmployeeUpdate.metroComboBox1.DisplayMember = "Зарплата";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, int> list = new Dictionary<int, int>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (int)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            con.Close();
            ObjEmployeeUpdate.metroComboBox1.DataSource = list.ToList();
            ObjEmployeeUpdate.metroComboBox1.DisplayMember = "Value";
            ObjEmployeeUpdate.metroComboBox1.ValueMember = "Key";
            ObjEmployeeUpdate.metroComboBox1.SelectedValue = EMPLmetroGrid1.CurrentRow.Cells[10].Value;
            if (ObjEmployeeUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    con.Close();
                    EMPLmetroGrid1.Sort(EMPLmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update Сотрудник set Фамилия=@st1, Имя=@st2, Отчество=@st3, Должность=@st4,Телефон=@st5, Дата=@st6,Пароль=@st7,Фото=@st8,Идзарплата=@st9 where идсотрудник=" + numbstrokemployee + "", con);
                    sss.Parameters.AddWithValue("st1", ObjEmployeeUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjEmployeeUpdate.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjEmployeeUpdate.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjEmployeeUpdate.metroTextBox4.Text);
                    sss.Parameters.AddWithValue("st5", ObjEmployeeUpdate.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjEmployeeUpdate.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st7", Convert.ToInt32(ObjEmployeeUpdate.metroTextBox5.Text));
                    sss.Parameters.AddWithValue("st8", ObjEmployeeUpdate.metroTextBox6.Text);
                    sss.Parameters.AddWithValue("st9", Convert.ToInt32(ObjEmployeeUpdate.metroComboBox1.SelectedValue));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updEmployee();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Close();
                con.Open();
                numbstrokemployee = Convert.ToInt32(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM сотрудник 
                                                    WHERE идсотрудник=" + numbstrokemployee + "", con);
                sss.ExecuteNonQuery();
                updEmployee();
                con.Close();
            }
        }

        private void metroTile8_Click(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updEmployee();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Сотрудники";
            ExcelApp.Cells[1, 1] = "Фамилия";
            ExcelApp.Cells[1, 2] = "Имя";
            ExcelApp.Cells[1, 3] = "Отчество";
            ExcelApp.Cells[1, 4] = "Должность";
            ExcelApp.Cells[1, 5] = "Телефон";
            ExcelApp.Cells[1, 6] = "Дата";
            ExcelApp.Cells[1, 7] = "Пароль";
            ExcelApp.Cells[1, 8] = "Зарплата";
            {
                for (int i = 1; i < EMPLmetroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = EMPLmetroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = EMPLmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = EMPLmetroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = EMPLmetroGrid1.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = EMPLmetroGrid1.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = EMPLmetroGrid1.Rows[i - 1].Cells[6].Value;
                    ExcelApp.Cells[i + 1, 7] = EMPLmetroGrid1.Rows[i - 1].Cells[7].Value;
                    ExcelApp.Cells[i + 1, 8] = EMPLmetroGrid1.Rows[i - 1].Cells[9].Value;
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile9_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Employee.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, EMPLmetroGrid1.RowCount + 1, 8);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя"; ;
                table.Cell(1, 3).Range.Text = "Отчество";
                table.Cell(1, 4).Range.Text = "Должность";
                table.Cell(1, 5).Range.Text = "Телефон";
                table.Cell(1, 6).Range.Text = "Дата";
                table.Cell(1, 7).Range.Text = "Пароль";
                table.Cell(1, 8).Range.Text = "Зарплата";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < EMPLmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[6].Value.ToString();
                    table.Cell(i + 1, 7).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[7].Value.ToString();
                    table.Cell(i + 1, 8).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[9].Value.ToString();
                }

                try
                {
                    doc.SaveAs(path);
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch { }
            }
            catch { }
        }

        private void metroTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            //con.Close();
            EMPLmetroTabControl5.Enabled = false;
            EMPLmetroTile11.Enabled = false;
            EMPLmetroTabControl7.Enabled = false;
            EMPLmetroTabControl8.Enabled = false;
            EMPLmetroTile16.Enabled = false;
            EMPLmetroTile3.Visible = true;
            string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                       FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                       WHERE[Фамилия] LIKE '%" + EMPLtextBox2.Text + "%'";
            sdaEmployee = new OleDbDataAdapter(s, con);
            dtEmployee = new DataTable();
            sdaEmployee.Fill(dtEmployee);
            EMPLmetroGrid1.DataSource = dtEmployee;
            if (EMPLmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                EMPLtextBox2.Text = "";
                updEmployee();
            }
        }

        public void EnableOFF()
        {
            EMPLmetroTabControl3.Enabled = false;
            EMPLmetroTabControl4.Enabled = false;
            EMPLmetroTabControl5.Enabled = false;
            EMPLmetroTabControl6.Enabled = false;
            EMPLmetroTabControl7.Enabled = false;
            EMPLmetroTabControl8.Enabled = false;
            EMPLmetroTile3.Visible = false;
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            EMPLmetroTile3.Visible = true;
            //string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            OleDbConnection conn = new OleDbConnection(conString);
          //  conn.Close();
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = EMPLmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = EMPLmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(EMPLmetroDateTime1.Text) < Convert.ToDateTime(EMPLmetroDateTime2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
WHERE сотрудник.Дата Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                EMPLmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                EMPLmetroTabControl5.Enabled = false;
                EMPLmetroTabControl7.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                EMPLmetroTile10.Enabled = false;
                EMPLmetroTile16.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                //EnableOFF();
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updEmployee();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updEmployee();
            }
        }

        private void metroTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox3_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox4_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox5_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroTextBox8_Click(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            if (EMPLtextBox3.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника 
                           INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                           WHERE Фамилия='" + EMPLtextBox3.Text + "'";
                sdaEmployee = new OleDbDataAdapter(s, con);
                dtEmployee = new DataTable();
                sdaEmployee.Fill(dtEmployee);
                EMPLmetroGrid1.DataSource = dtEmployee;
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroGrid1.Columns[0].Visible = false;
                EMPLmetroTabControl5.Enabled = false;
                EMPLmetroTabControl6.Enabled = false;
                EMPLmetroTile14.Enabled = false;
                EMPLmetroTile15.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                EMPLmetroTile3.Visible = true;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updEmployee();
                }
                EMPLmetroGrid1.Text = "";
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if (EMPLtextBox4.Text == "")
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {

                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Имя='" + EMPLtextBox4.Text + "'";
                sdaEmployee = new OleDbDataAdapter(s, con);
                dtEmployee = new DataTable();
                sdaEmployee.Fill(dtEmployee);
                EMPLmetroGrid1.DataSource = dtEmployee;
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroGrid1.Columns[0].Visible = false;
                EMPLmetroTabControl5.Enabled = false;
                EMPLmetroTabControl3.Enabled = false;
                EMPLmetroTile14.Enabled = false;
                EMPLmetroTile15.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                EMPLmetroTile3.Visible = true;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updEmployee();
                    EMPLmetroTile3.Visible = true;
                }
                EMPLtextBox4.Text = "";
            }
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            EMPLmetroTile3.Visible = true;
            if ((EMPLtextBox5.Text == "") || (EMPLtextBox8.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Фамилия='" + EMPLtextBox5.Text + "' and Имя='" + EMPLtextBox8.Text + "'";
                sdaEmployee = new OleDbDataAdapter(s, con);
                dtEmployee = new DataTable();
                sdaEmployee.Fill(dtEmployee);
                EMPLmetroGrid1.DataSource = dtEmployee;
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroTabControl4.Enabled = false;
                EMPLmetroTabControl3.Enabled = false;
                EMPLmetroTile14.Enabled = false;
                EMPLmetroTile15.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                EMPLmetroTile3.Visible = true;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updEmployee();
                }
                EMPLtextBox5.Text = "";
                EMPLtextBox8.Text = "";
            }
        }

        private void сброситьФильтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updEmployee();
        }

        private void metroGrid1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (EMPLmetroGrid1.CurrentRow != null)
                {
                    EMPLpictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    metroTextBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[8].Value);
                    EMPLpictureBox1.Load(Application.StartupPath + @"\PhotoEmployee\" + metroTextBox1.Text);
                }
            }
           catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
            }
        }
        

        private void metroTile12_Click(object sender, EventArgs e)
        {
            
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@" select Фамилия, Имя, Отчество, Должность, Телефон, Дата,Пароль,Идзарплата ,Фото
                                                                              FROM сотрудник 
                                                                              WHERE (сотрудник.[Дата] Is Not Null) 
                                                                              AND ((DateDiff('y',Date(),DateAdd('yyyy',DateDiff('yyyy',[Дата],Date()),[Дата])) 
                                                                              Between 0 And 31))"), conn);
            adapter.Fill(ds);
            EMPLmetroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            if (EMPLmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updEmployee();
            }
        }

        private void metroTile17_Click(object sender, EventArgs e)
        {
          
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT TOP 1 Сотрудник.Фамилия , Сотрудник.Имя , Сотрудник.Отчество, Сотрудник.Должность,Сотрудник.Телефон,Сотрудник.Дата,Сотрудник.Пароль, Count(Продажа_абонемента.Идпродажа) AS [Количество продаж] ,Сотрудник.фото
                                                                         FROM Сотрудник 
                                                                         INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                         ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник 
                                                                         GROUP BY Сотрудник.Фамилия , Сотрудник.Имя , Сотрудник.Отчество, Сотрудник.Должность,Сотрудник.Телефон,Сотрудник.Дата,Сотрудник.Пароль, Сотрудник.Идсотрудник,сотрудник.фото  
                                                                         ORDER BY Count(Продажа_абонемента.Идпродажа) desc;"), conn);
            adapter.Fill(ds);
            EMPLmetroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            EMPLmetroTabControl3.Enabled = false;
            if (EMPLmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);
                updEmployee();
            }
        }

        private void metroTile19_Click(object sender, EventArgs e)
        {
         
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Абонемент.Название AS Абонемент, Сотрудник.Фамилия  & Сотрудник.Имя AS [Реализовавший сотрудник], Спортсмен.Фамилия  & Спортсмен.Имя AS Спортсмен, Продажа_абонемента.Дата_начала as Начало,  Продажа_абонемента.Дата_окончания as Конец,сотрудник.фото,сотрудник.фото,сотрудник.фото,сотрудник.фото 
                                                                          FROM Спортсмен 
                                                                          INNER JOIN (Сотрудник 
                                                                          INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                          ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник) 
                                                                          ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;"), conn);
            adapter.Fill(ds);
            EMPLmetroGrid1.DataSource = ds.Tables[0];
            EMPLmetroGrid1.Columns[8].Visible = false;
            EMPLmetroGrid1.Columns[7].Visible = false;
            EMPLmetroGrid1.Columns[6].Visible = false;
            EMPLmetroGrid1.Columns[5].Visible = false;
            conn.Close();
            EnableOFF();
            EMPLmetroTabControl3.Enabled = false;
            if (EMPLmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updEmployee();
            }
        }

        private void metroTile7_Click(object sender, EventArgs e)
        {
            Money main2 = new Money();
            main2.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            EMPLmetroTile3.Visible = true;
    
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            if (Convert.ToInt32(EMPLmetroTextBox6.Text) < Convert.ToInt32(EMPLmetroTextBox7.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата where зарплата_сотрудника.зарплата between {0} and {1}", EMPLmetroTextBox6.Text, EMPLmetroTextBox7.Text), conn);
                adapter.Fill(ds);
                EMPLmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                EMPLmetroTile10.Enabled = false;
                EMPLmetroTile11.Enabled = false;
                EMPLmetroTabControl5.Enabled = false;
                EMPLmetroTabControl7.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updEmployee();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальная зарплата не может быть больше конечной", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updEmployee();
            }
        }

        private void metroTile21_Click(object sender, EventArgs e)
        {
            Send_Mail send = new Send_Mail();
            send.ShowDialog();
        }

        private void metroTile23_Click(object sender, EventArgs e)
        {
            Report_Employee rep = new Report_Employee();
            rep.Show();
        }
        /////////////////////////////////////////////        спортсмен          //////////////////////////////////////////////////////////////////////////////

        private void metroTile41_Click(object sender, EventArgs e)
        {
            MOD_sportsmen ObjSportsmenAdd = new MOD_sportsmen();
            ObjSportsmenAdd.textBox1.Text = "";
            ObjSportsmenAdd.textBox2.Text = "";
            ObjSportsmenAdd.textBox3.Text = "";
            ObjSportsmenAdd.maskedTextBox1.Text = "";
            ObjSportsmenAdd.metroDateTime1.Text = null;
            // ObjTrenerAdd.metroComboBox1.Text =;
            ObjSportsmenAdd.Text = "Добавить спортсмена";
            ObjSportsmenAdd.metroTile1.Text = "Добавить";

            if (ObjSportsmenAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    SPORTMmetroGrid2.Sort(SPORTMmetroGrid2.Columns[1], ListSortDirection.Ascending);
                    idsportsmen = Convert.ToString(Convert.ToInt32(SPORTMmetroGrid2.Rows[SPORTMmetroGrid2.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [спортсмен]
                                                        ( Фамилия, Имя, Отчество, Телефон, Дата_рождения, Пол)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5,@st6)", con);
                    sss.Parameters.AddWithValue("st1", ObjSportsmenAdd.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjSportsmenAdd.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjSportsmenAdd.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjSportsmenAdd.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st5", Convert.ToDateTime(ObjSportsmenAdd.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st6", ObjSportsmenAdd.metroComboBox1.Text);
                    sss.ExecuteNonQuery();
                    con.Close();
                    updSportsmen();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile40_Click(object sender, EventArgs e)
        {
            MOD_sportsmen ObjSportsmenUpdate = new MOD_sportsmen();

            ObjSportsmenUpdate.Text = "Редактировать спортсмена";
            ObjSportsmenUpdate.metroTile1.Text = "Редактировать";
            numbstrokasportsmen = Convert.ToInt32(SPORTMmetroGrid2.CurrentRow.Cells[0].Value);
            ObjSportsmenUpdate.textBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[1].Value);
            ObjSportsmenUpdate.textBox2.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[2].Value);
            ObjSportsmenUpdate.textBox3.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[3].Value);
            ObjSportsmenUpdate.maskedTextBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[4].Value);
            ObjSportsmenUpdate.metroDateTime1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[5].Value);
            ObjSportsmenUpdate.metroComboBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[6].Value);
            if (ObjSportsmenUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    SPORTMmetroGrid2.Sort(SPORTMmetroGrid2.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update спортсмен set Фамилия=@st1, Имя=@st2, Отчество=@st3, Телефон=@st4, Дата_рождения=@st5,Пол=@st6 where идспортсмен=" + numbstrokasportsmen + "", con);
                    sss.Parameters.AddWithValue("st1", ObjSportsmenUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjSportsmenUpdate.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjSportsmenUpdate.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjSportsmenUpdate.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st5", Convert.ToDateTime(ObjSportsmenUpdate.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st6", ObjSportsmenUpdate.metroComboBox1.Text);
                    sss.ExecuteNonQuery();
                    con.Close();
                    updSportsmen();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile39_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Open();
                numbstrokasportsmen = Convert.ToInt32(SPORTMmetroGrid2.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM спортсмен 
                                                    WHERE идспортсмен=" + numbstrokasportsmen + "", con);
                sss.ExecuteNonQuery();
                updSportsmen();
                con.Close();
            }
        }

        private void metroTile37_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updEmployee();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Спортсмен";
            ExcelApp.Cells[1, 1] = "Фамилия";
            ExcelApp.Cells[1, 2] = "Имя";
            ExcelApp.Cells[1, 3] = "Отчество";
            ExcelApp.Cells[1, 4] = "Телефон";
            ExcelApp.Cells[1, 5] = "Дата";
            ExcelApp.Cells[1, 6] = "Пол";
            {
                for (int i = 1; i < SPORTMmetroGrid2.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = SPORTMmetroGrid2.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = SPORTMmetroGrid2.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = SPORTMmetroGrid2.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = SPORTMmetroGrid2.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = SPORTMmetroGrid2.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = SPORTMmetroGrid2.Rows[i - 1].Cells[6].Value;

                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }


        private void metroTile36_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path1 = Directory.GetCurrentDirectory() + @"\" + "report/Sportsmen.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, SPORTMmetroGrid2.RowCount + 1, 6);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя"; ;
                table.Cell(1, 3).Range.Text = "Отчество";

                table.Cell(1, 4).Range.Text = "Телефон";
                table.Cell(1, 5).Range.Text = "Дата";
                table.Cell(1, 6).Range.Text = "Пол";

                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < SPORTMmetroGrid2.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = SPORTMmetroGrid2.Rows[i - 1].Cells[6].Value.ToString();
                }

                try
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path1, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    doc.SaveAs(path1);

                }
                catch { }
            }
            catch { }


            ///////////////////////////////////////////////////////////////////////
        }
        
        private void HeadForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Exit s = new Exit();
            if (s.ShowDialog() == DialogResult.OK)
            {
                Environment.Exit(0);
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox10_KeyUp(object sender, KeyEventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            metroTile33.Enabled = false;
            metroTile32.Enabled = false;
            string s = @"SELECT *
                       FROM спортсмен 
                       WHERE[Фамилия] LIKE '%" + SPORTMtextBox10.Text + "%'";
            sdasportsmen = new OleDbDataAdapter(s, con);
            dtSportsmen = new DataTable();
            sdasportsmen.Fill(dtSportsmen);
            SPORTMmetroGrid2.DataSource = dtSportsmen;
            if (SPORTMmetroGrid2.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SPORTMtextBox10.Text = "";
                updSportsmen();
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            metroTile34.Enabled = false;
            metroTile32.Enabled = false;
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = SPORTMmetroDateTime3.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = SPORTMmetroDateTime4.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(SPORTMmetroDateTime3.Text) > Convert.ToDateTime(SPORTMmetroDateTime4.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT *
                                                                              FROM спортсмен
                                                                              WHERE Дата_рождения Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                SPORTMmetroGrid2.DataSource = ds.Tables[0];
                conn.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSportsmen();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updSportsmen();
            }
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            switch (SPORTMmetroComboBox1.SelectedIndex)
            {

                case 0:
                    SPORTMmetroTabControl12.Enabled = false;
                    SPORTMmetroTabControl10.Enabled = false;
                    SPORTMmetroTabControl9.Enabled = false;
                    metroTile34.Enabled = false;
                    metroTile33.Enabled = false;
                    string s = @"SELECT * FROM Спортсмен WHERE (((Спортсмен.Пол) = 'Мужской'));";
                    sdasportsmen = new OleDbDataAdapter(s, con);
                    dtSportsmen = new DataTable();
                    sdasportsmen.Fill(dtSportsmen);
                    SPORTMmetroGrid2.DataSource = dtSportsmen;
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        updSportsmen();
                    }
                    break;
                case 1:
                    SPORTMmetroTabControl12.Enabled = false;
                    SPORTMmetroTabControl10.Enabled = false;
                    SPORTMmetroTabControl9.Enabled = false;
                    metroTile34.Enabled = false;
                    metroTile33.Enabled = false;
                    string s1 = @"SELECT * FROM Спортсмен WHERE (((Спортсмен.Пол) = 'Женский'));";
                    sdasportsmen = new OleDbDataAdapter(s1, con);
                    dtSportsmen = new DataTable();
                    sdasportsmen.Fill(dtSportsmen);
                    SPORTMmetroGrid2.DataSource = dtSportsmen;
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсменов с таким полом не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        updSportsmen();
                    }
                    break;


            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updSportsmen();
        }

        private void metroComboBox1_Click(object sender, EventArgs e)
        {
           
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            if (SPORTMtextBox9.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl11.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                metroTile30.Enabled = false;
                metroTile29.Enabled = false;
                string s = @"SELECT *
                           FROM спортсмен 
                           WHERE Фамилия='" + SPORTMtextBox9.Text + "'";
                sdasportsmen = new OleDbDataAdapter(s, con);
                dtSportsmen = new DataTable();
                sdasportsmen.Fill(dtSportsmen);
                SPORTMmetroGrid2.DataSource = dtSportsmen;
                //metroGrid2.Columns[8].Visible = false;
                SPORTMmetroGrid2.Columns[0].Visible = false;
               
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSportsmen();
                }
                SPORTMtextBox9.Text = "";
            }
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {
            if (SPORTMtextBox7.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl11.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                metroTile31.Enabled = false;
                metroTile29.Enabled = false;
                string s = @"SELECT *
                           FROM спортсмен 
                           WHERE Имя='" + SPORTMtextBox7.Text + "'";
                sdasportsmen = new OleDbDataAdapter(s, con);
                dtSportsmen = new DataTable();
                sdasportsmen.Fill(dtSportsmen);
                SPORTMmetroGrid2.DataSource = dtSportsmen;
                //metroGrid2.Columns[8].Visible = false;
                SPORTMmetroGrid2.Columns[0].Visible = false;

                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSportsmen();
                }
                SPORTMtextBox7.Text = "";
            }
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            //EMPLmetroTile3.Visible = true;
            if ((SPORTMtextBox6.Text == "") || (SPORTMtextBox1.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl11.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                metroTile31.Enabled = false;
                metroTile20.Enabled = false;
                string s = @"SELECT *
                           from спортсмен
                           WHERE Фамилия='" + SPORTMtextBox6.Text + "' and Имя='" + SPORTMtextBox1.Text + "'";
                sdasportsmen = new OleDbDataAdapter(s, con);
                dtSportsmen = new DataTable();
                sdasportsmen.Fill(dtSportsmen);
                SPORTMmetroGrid2.DataSource = dtSportsmen;
                SPORTMmetroGrid2.Columns[0].Visible = false;
            
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSportsmen();
                }
                SPORTMtextBox6.Text = "";
                SPORTMtextBox1.Text = "";
            }
        }

        private void metroTile28_Click(object sender, EventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl11.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl14.Enabled = false;
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@" SELECT Спортсмен.Фамилия,Спортсмен.Имя,Спортсмен.Отчество, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Абонемент.Название, Абонемент.Количество_посещений
                                                                          FROM Абонемент 
                                                                          INNER JOIN (Спортсмен
                                                                          INNER JOIN Продажа_абонемента
                                                                          ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) 
                                                                          ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент;"), conn);
            adapter.Fill(ds);
            SPORTMmetroGrid2.DataSource = ds.Tables[0];
            conn.Close();
            if (SPORTMmetroGrid2.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updSportsmen();
            }
        }

        private void metroTile27_Click(object sender, EventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl11.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl14.Enabled = false;
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            try
            {
                string n = Microsoft.VisualBasic.Interaction.InputBox("Введите Фамилию спортсмена: ", "Запрос",
                    "", 500, 500);
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@" SELECT Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество, Учет_посещений.Дата,Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания
FROM (Спортсмен INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) INNER JOIN Учет_посещений ON Продажа_абонемента.Идпродажа = Учет_посещений.Идпродажа
WHERE (((Спортсмен.Фамилия)='" + n + "'));"), conn);

                adapter.Fill(ds);
                SPORTMmetroGrid2.DataSource = ds.Tables[0];
                conn.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    updSportsmen();
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
            }
        }

        private void metroTile26_Click(object sender, EventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl11.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl14.Enabled = false;
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
           
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@" SELECT Идспортсмен, Фамилия, Имя, Отчество,
       (SELECT COUNT(*) FROM Учет_посещений WHERE Идпродажа IN 
        (SELECT Идпродажа FROM продажа_абонемента WHERE Идспортсмен = Спортсмен.Идспортсмен)) AS КоличествоПосещений
               FROM Спортсмен"), conn);
            adapter.Fill(ds);
            SPORTMmetroGrid2.DataSource = ds.Tables[0];
            conn.Close();
            if (SPORTMmetroGrid2.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updSportsmen();
            }
        }

        private void metroTile19_Click_1(object sender, EventArgs e)
        {
            MOD_Trener ObjTrenerAdd = new MOD_Trener();
            ObjTrenerAdd.textBox1.Text = "";
            ObjTrenerAdd.textBox2.Text = "";
            ObjTrenerAdd.textBox3.Text = "";
            ObjTrenerAdd.maskedTextBox1.Text = "";
            ObjTrenerAdd.metroDateTime1.Text = null;
           // ObjTrenerAdd.metroTextBox6.Text = "";
            ObjTrenerAdd.metroTextBox5.Text = "";
            ObjTrenerAdd.Text = "Добавить тренера";
            ObjTrenerAdd.metroTile1.Text = "Добавить";
            ObjTrenerAdd.metroTextBox1.Text = "";
            if (ObjTrenerAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENmetroGrid1.Sort(TRENmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    idtrener = Convert.ToString(Convert.ToInt32(TRENmetroGrid1.Rows[TRENmetroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [тренер]
                                                        ( Фамилия, Имя, Отчество, Должность,Телефон, Дата_рождения,Пароль,Фото,оклад)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5,@st6,@st7,@st8,@st9)", con);
                    sss.Parameters.AddWithValue("st1", ObjTrenerAdd.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjTrenerAdd.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjTrenerAdd.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjTrenerAdd.metroTextBox4.Text);
                    sss.Parameters.AddWithValue("st5", ObjTrenerAdd.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjTrenerAdd.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st7", Convert.ToInt32(ObjTrenerAdd.metroTextBox5.Text));
                    sss.Parameters.AddWithValue("st8", ObjTrenerAdd.metroTextBox6.Text);
                    sss.Parameters.AddWithValue("st9", Convert.ToInt32(ObjTrenerAdd.metroTextBox1.Text));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updTrener();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile18_Click_1(object sender, EventArgs e)
        {
            MOD_Trener ObjTrenerUpdate = new MOD_Trener();
            ObjTrenerUpdate.Text = "Редактировать тренера";
            ObjTrenerUpdate.metroTile1.Text = "Редактировать";
            numbstroktrener = Convert.ToInt32(TRENmetroGrid1.CurrentRow.Cells[0].Value);
            ObjTrenerUpdate.textBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[1].Value);
            ObjTrenerUpdate.textBox2.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[2].Value);
            ObjTrenerUpdate.textBox3.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[3].Value);
            ObjTrenerUpdate.metroTextBox4.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[4].Value);
            ObjTrenerUpdate.maskedTextBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[5].Value);
            ObjTrenerUpdate.metroDateTime1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[6].Value);
            ObjTrenerUpdate.metroTextBox5.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[7].Value);
            ObjTrenerUpdate.metroTextBox6.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[8].Value);
            ObjTrenerUpdate.metroTextBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[9].Value);
            if (ObjTrenerUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENmetroGrid1.Sort(TRENmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update тренер set Фамилия=@st1, Имя=@st2, Отчество=@st3, Должность=@st4,Телефон=@st5, Дата_рождения=@st6,Пароль=@st7,Фото=@st8,оклад=@st9 where идтренер=" + numbstroktrener + "", con);
                    sss.Parameters.AddWithValue("st1", ObjTrenerUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjTrenerUpdate.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjTrenerUpdate.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjTrenerUpdate.metroTextBox4.Text);
                    sss.Parameters.AddWithValue("st5", ObjTrenerUpdate.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjTrenerUpdate.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st7", Convert.ToInt32(ObjTrenerUpdate.metroTextBox5.Text));
                    sss.Parameters.AddWithValue("st8", ObjTrenerUpdate.metroTextBox6.Text);
                    sss.Parameters.AddWithValue("st9", Convert.ToInt32(ObjTrenerUpdate.metroTextBox1.Text));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updTrener();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile17_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Open();
                numbstroktrener = Convert.ToInt32(TRENmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM тренер 
                                                    WHERE идтренер=" + numbstroktrener + "", con);
                sss.ExecuteNonQuery();
                updTrener();
                con.Close();
            }
        }

        private void metroTile15_Click(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updEmployee();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Тренер";
            ExcelApp.Cells[1, 1] = "Фамилия";
            ExcelApp.Cells[1, 2] = "Имя";
            ExcelApp.Cells[1, 3] = "Отчество";
            ExcelApp.Cells[1, 4] = "Должность";
            ExcelApp.Cells[1, 5] = "Телефон";
            ExcelApp.Cells[1, 6] = "Дата";
            ExcelApp.Cells[1, 7] = "Пароль";
            ExcelApp.Cells[1, 8] = "Оклад";
            {
                for (int i = 1; i < TRENmetroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = TRENmetroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = TRENmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = TRENmetroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = TRENmetroGrid1.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = TRENmetroGrid1.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = TRENmetroGrid1.Rows[i - 1].Cells[6].Value;
                    ExcelApp.Cells[i + 1, 7] = TRENmetroGrid1.Rows[i - 1].Cells[7].Value;
                    ExcelApp.Cells[i + 1, 8] = TRENmetroGrid1.Rows[i - 1].Cells[9].Value;
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile14_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trener.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, TRENmetroGrid1.RowCount + 1, 8);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя"; ;
                table.Cell(1, 3).Range.Text = "Отчество";
                table.Cell(1, 4).Range.Text = "Должность";
                table.Cell(1, 5).Range.Text = "Телефон";
                table.Cell(1, 6).Range.Text = "Дата";
                table.Cell(1, 7).Range.Text = "Пароль";
                table.Cell(1, 8).Range.Text = "Оклад";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < TRENmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[6].Value.ToString();
                    table.Cell(i + 1, 7).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[7].Value.ToString();
                    table.Cell(i + 1, 8).Range.Text = TRENmetroGrid1.Rows[i - 1].Cells[9].Value.ToString();
                }

                try
                {
                    doc.SaveAs(path);
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch { }
            }
            catch { }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox8_KeyUp(object sender, KeyEventArgs e)
        {
            string s = @"SELECT Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                       FROM тренер
                       WHERE[Фамилия] LIKE '%" + TRENtextBox8.Text + "%'";
            sdatrener = new OleDbDataAdapter(s, con);
            dtTrener = new DataTable();
            sdatrener.Fill(dtTrener);
            TRENmetroGrid1.DataSource = dtTrener;
            metroTile11.Enabled=false;
            metroTile10.Enabled = false;
            TRENmetroTabControl4.Enabled = false;
            TRENmetroTabControl3.Enabled = false;
            TRENmetroTabControl6.Enabled = false;
            if (TRENmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                TRENtextBox8.Text = "";
                updTrener();
            }
        }

        private void metroButton5_Click_1(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = TRENmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = TRENmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(TRENmetroDateTime2.Text) < Convert.ToDateTime(TRENmetroDateTime1.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter
                (String.Format(@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                                WHERE Дата_рождения Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                TRENmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTile12.Enabled = false;
                metroTile10.Enabled = false;
                TRENmetroTabControl4.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;

                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updTrener();
            }
        }

        private void metroButton4_Click_1(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            if (Convert.ToInt32(TRENmetroTextBox3.Text) < Convert.ToInt32(TRENmetroTextBox2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                               where оклад between {0} and {1}", TRENmetroTextBox3.Text, TRENmetroTextBox2.Text), conn);
                adapter.Fill(ds);
                TRENmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTile12.Enabled = false;
                metroTile11.Enabled = false;
                TRENmetroTabControl4.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный оклад не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updTrener(); 
            }
        }

        private void metroButton3_Click_1(object sender, EventArgs e)
        {
            if (TRENtextBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Фамилия='" + TRENtextBox5.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                TRENmetroGrid1.DataSource = dtTrener;
                TRENmetroGrid1.Columns[8].Visible = false;
                //  metroGrid1.Columns[0].Visible = false;
                metroTile7.Enabled = false;
                metroTile8.Enabled = false;
                TRENmetroTabControl5.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                TRENtextBox5.Text = "";
            }
        }

        private void metroButton2_Click_1(object sender, EventArgs e)
        {
            if (TRENtextBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Имя='" + TRENtextBox4.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                TRENmetroGrid1.DataSource = dtTrener;
                TRENmetroGrid1.Columns[8].Visible = false;
               // metroGrid1.Columns[0].Visible = false;
                metroTile7.Enabled = false;
                metroTile9.Enabled = false;
                TRENmetroTabControl5.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;

                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                TRENtextBox4.Text = "";
            }
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
           // EMPLmetroTile3.Visible = true;
            if ((TRENtextBox3.Text == "") || (TRENtextBox2.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string s = @"SELECT Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                               WHERE Фамилия='" + TRENtextBox3.Text + "' and Имя='" + TRENtextBox2.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                TRENmetroGrid1.DataSource = dtTrener;
                TRENmetroGrid1.Columns[8].Visible = false;
                metroTile8.Enabled = false;
                metroTile9.Enabled = false;
                TRENmetroTabControl5.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                TRENtextBox3.Text = "";
                TRENtextBox2.Text = "";
            }
        }

        private void metroGrid1_SelectionChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (TRENmetroGrid1.CurrentRow != null)
                {
                    TRENpictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                    textBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[8].Value);
                    TRENpictureBox2.Load(Application.StartupPath + @"\PhotoTrener\" + textBox1.Text);
                }
            }
            catch (Exception ex)
            {
                MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
            }
        }

        private void сброситьФильтрToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updTrener();
        }

        private void metroTile6_Click_1(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(conString);
                conn.Open();
                DataSet ds = new DataSet();
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Тренер.Фамилия, Тренер.Имя, Тренер.Отчество, Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество,Тренировка.Название,тренер.Телефон,Тренер.Фото
                                                                              FROM Спортсмен 
                                                                              INNER JOIN (((Тренер INNER JOIN Тренировка ON Тренер.Идтренер = Тренировка.Идтренер)
                                                                              INNER JOIN Абонемент ON Тренировка.Идтренировка = Абонемент.Идтренеровка) 
                                                                              INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                              ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;"), conn);

                metroTile16.Visible = false;
                TRENpictureBox2.Visible=false;

                adapter.Fill(ds);
                TRENmetroGrid1.DataSource = ds.Tables[0];

                TRENmetroTabControl8.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl4.Enabled = false;
                TRENmetroTabControl5.Enabled = false;
                conn.Close();
                TRENmetroGrid1.Columns[8].Visible = false;


                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    updTrener();
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void metroTile13_Click(object sender, EventArgs e)
        {
            Report_Trener reptr = new Report_Trener();
            reptr.Show();
                
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            RepSportsmen resp = new RepSportsmen();
            resp.Show();
        }

        private void metroTile38_Click(object sender, EventArgs e)
        {
            MOD_Trenerovka ObjTrenerovkaAdd = new MOD_Trenerovka();
        
            ObjTrenerovkaAdd.textBox1.Text = "";
            ObjTrenerovkaAdd.Text = "Добавить тренировку";
            ObjTrenerovkaAdd.metroTile1.Text = "Добавить";
            con.Open();
            OleDbCommand cmd = new OleDbCommand(@"SELECT Вид_тренировки.Идвидтренировка, Вид_тренировки.Название FROM Вид_тренировки;", con);
            ObjTrenerovkaAdd.metroComboBox1.DisplayMember = "Название";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjTrenerovkaAdd.metroComboBox1.DataSource = list.ToList();
            ObjTrenerovkaAdd.metroComboBox1.DisplayMember = "Value";
            ObjTrenerovkaAdd.metroComboBox1.ValueMember = "Key";
            con.Close();

            con.Open();
            OleDbCommand cmd1 = new OleDbCommand(@"SELECT Тренер.Идтренер, Тренер.Фамилия FROM Тренер;", con);
            ObjTrenerovkaAdd.metroComboBox2.DisplayMember = "Фамилия";
            OleDbDataReader reader1 = cmd1.ExecuteReader();
            Dictionary<int, string> list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int)reader1[0], (string)reader1[1]);
            }
            reader1.Close();
            cmd1.ExecuteNonQuery();
            ObjTrenerovkaAdd.metroComboBox2.DataSource = list1.ToList();
            ObjTrenerovkaAdd.metroComboBox2.DisplayMember = "Value";
            ObjTrenerovkaAdd.metroComboBox2.ValueMember = "Key";
            con.Close();
         //   ShowDialog();
            if (ObjTrenerovkaAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENINGmetroGrid1.Sort(TRENINGmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    idtrenerovka = Convert.ToString(Convert.ToInt32(TRENINGmetroGrid1.Rows[TRENINGmetroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [тренировка]
                                                        ( Название, Идвидтренировка, Идтренер)
                                                        VALUES(@st1,@st2,@st3)", con);
                    sss.Parameters.AddWithValue("st1", ObjTrenerovkaAdd.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjTrenerovkaAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.Parameters.AddWithValue("st3", Convert.ToInt32(ObjTrenerovkaAdd.metroComboBox2.SelectedValue.ToString()));
                  
                    sss.ExecuteNonQuery();
                    con.Close();
                    updTrenerovka();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile37_Click_1(object sender, EventArgs e)
        {
            MOD_Trenerovka ObjTrenerovkaUpdate = new MOD_Trenerovka();
            con.Close();
            ObjTrenerovkaUpdate.Text = "Редактировать тренировку";
            ObjTrenerovkaUpdate.metroTile1.Text = "Редактировать";
            numbstrokTrenerovka = Convert.ToInt32(TRENINGmetroGrid1.CurrentRow.Cells[0].Value);
            ObjTrenerovkaUpdate.textBox1.Text = Convert.ToString(TRENINGmetroGrid1.CurrentRow.Cells[1].Value);
            con.Open();
            OleDbCommand cmd = new OleDbCommand(@"SELECT Вид_тренировки.Идвидтренировка, Вид_тренировки.Название FROM Вид_тренировки;", con);
            ObjTrenerovkaUpdate.metroComboBox1.DisplayMember = "Название";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            
            ObjTrenerovkaUpdate.metroComboBox1.DataSource = list.ToList();
            ObjTrenerovkaUpdate.metroComboBox1.DisplayMember = "Value";
            ObjTrenerovkaUpdate.metroComboBox1.ValueMember = "Key";
           
            ObjTrenerovkaUpdate.metroComboBox1.SelectedValue = TRENINGmetroGrid1.CurrentRow.Cells[5].Value;

            con.Close();
            con.Open();
            OleDbCommand cmd1 = new OleDbCommand(@"SELECT Тренер.Идтренер, Тренер.Фамилия FROM Тренер;", con);
            ObjTrenerovkaUpdate.metroComboBox2.DisplayMember = "Фамилия";
            OleDbDataReader reader1 = cmd1.ExecuteReader();
            Dictionary<int, string> list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int)reader1[0], (string)reader1[1]);
            }
            reader1.Close();
            cmd1.ExecuteNonQuery();
            
            ObjTrenerovkaUpdate.metroComboBox2.DataSource = list1.ToList();
            ObjTrenerovkaUpdate.metroComboBox2.DisplayMember = "Value";
            ObjTrenerovkaUpdate.metroComboBox2.ValueMember = "Key";
            
            ObjTrenerovkaUpdate.metroComboBox2.SelectedValue = TRENINGmetroGrid1.CurrentRow.Cells[4].Value;
            con.Close();
            if (ObjTrenerovkaUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    con.Close();
                    TRENINGmetroGrid1.Sort(TRENINGmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update тренировка set название=@st1, Идвидтренировка=@st2, Идтренер=@st3 where Идтренировка=" + numbstrokTrenerovka + "", con);
                    sss.Parameters.AddWithValue("st1", ObjTrenerovkaUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjTrenerovkaUpdate.metroComboBox1.SelectedValue.ToString()));
                    sss.Parameters.AddWithValue("st3", Convert.ToInt32(ObjTrenerovkaUpdate.metroComboBox2.SelectedValue.ToString()));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updTrenerovka();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile36_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Close();
                con.Open();
                numbstrokTrenerovka = Convert.ToInt32(TRENINGmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM тренировка 
                                                    WHERE идтренировка=" + numbstrokTrenerovka + "", con);
                sss.ExecuteNonQuery();
                updTrenerovka();
                con.Close();
            }
        }

        private void metroTile28_Click_1(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updEmployee();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Сотрудники";
            ExcelApp.Cells[1, 1] = "Название";
            ExcelApp.Cells[1, 2] = "Вид";
            ExcelApp.Cells[1, 3] = "Тренер";
          
            {
                for (int i = 1; i < TRENINGmetroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = TRENINGmetroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = TRENINGmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = TRENINGmetroGrid1.Rows[i - 1].Cells[3].Value;
                   
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile27_Click_1(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trenerovka.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, TRENINGmetroGrid1.RowCount + 1, 3);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Навзание";
                table.Cell(1, 2).Range.Text = "Вид"; ;
                table.Cell(1, 3).Range.Text = "Тренер";
               
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < TRENINGmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                   
                }

                try
                {
                    doc.SaveAs(path);
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch { }
            }
            catch { }
        }

        private void textBox6_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox5_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyUp(object sender, KeyEventArgs e)
        {

            string s = @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер 
                       WHERE[Тренировка.Название] LIKE '%" + TRENINGtextBox6.Text + "%'";
            sdatrenerovka = new OleDbDataAdapter(s, con);
            dtTrenerovka = new DataTable();
            sdatrenerovka.Fill(dtTrenerovka);
            TRENINGmetroGrid1.DataSource = dtTrenerovka;
            TRENINGmetroTabControl6.Enabled = false;
            metroTile18.Enabled = false;
            TRENINGmetroTabControl4.Enabled = false;
            TRENINGmetroTabControl3.Enabled = false;
            if (TRENINGmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренировки не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                EMPLtextBox2.Text = "";
                updTrenerovka();
            }
        }

        private void metroButton3_Click_2(object sender, EventArgs e)
        {
            if (TRENINGtextBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
               
                string s = @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Вид_тренировки.Название='" + TRENINGtextBox5.Text + "'";
                sdatrenerovka = new OleDbDataAdapter(s, con);
                dtTrenerovka = new DataTable();
                sdatrenerovka.Fill(dtTrenerovka);
                TRENINGmetroGrid1.DataSource = dtTrenerovka;
                //metroGrid2.Columns[8].Visible = false;
                TRENINGmetroGrid1.Columns[0].Visible = false;
                TRENINGmetroGrid1.Columns[4].Visible = false;
                TRENINGmetroGrid1.Columns[5].Visible = false;
                TRENINGmetroTabControl3.Enabled = false;
                TRENINGmetroTabControl5.Enabled = false;
                TRENINGmetroTabControl6.Enabled = false;
                metroTile15.Enabled = false;
                metroTile14.Enabled = false;
                if (TRENINGmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких видов тренировок не найдено", "Вида тренировки не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrenerovka();
                }
                TRENINGtextBox5.Text = "";
            }
        }

        private void сброситьФильтрToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            updTrenerovka();
        }

        private void metroButton2_Click_2(object sender, EventArgs e)
        {

            if (TRENINGtextBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {

                string s = @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Тренер.Фамилия='" + TRENINGtextBox4.Text + "'";
                sdatrenerovka = new OleDbDataAdapter(s, con);
                dtTrenerovka = new DataTable();
                sdatrenerovka.Fill(dtTrenerovka);
                TRENINGmetroGrid1.DataSource = dtTrenerovka;
                //metroGrid2.Columns[8].Visible = false;
                TRENINGmetroGrid1.Columns[0].Visible = false;
                TRENINGmetroGrid1.Columns[4].Visible = false;
                TRENINGmetroGrid1.Columns[5].Visible = false;
                TRENINGmetroTabControl3.Enabled = false;
                TRENINGmetroTabControl5.Enabled = false;
                TRENINGmetroTabControl6.Enabled = false;
                metroTile17.Enabled = false;
                metroTile14.Enabled = false;
                if (TRENINGmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренеров не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrenerovka();
                }
                TRENINGtextBox4.Text = "";
            }
        }

        private void metroButton1_Click_2(object sender, EventArgs e)
        {
            if (TRENINGtextBox3.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {

                string s = @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Тренировка.Название='" + TRENINGtextBox3.Text + "'";
                sdatrenerovka = new OleDbDataAdapter(s, con);
                dtTrenerovka = new DataTable();
                sdatrenerovka.Fill(dtTrenerovka);
                TRENINGmetroGrid1.DataSource = dtTrenerovka;
                TRENINGmetroGrid1.Columns[0].Visible = false;
                TRENINGmetroGrid1.Columns[4].Visible = false;
                TRENINGmetroGrid1.Columns[5].Visible = false;
                TRENINGmetroTabControl3.Enabled = false;
                TRENINGmetroTabControl5.Enabled = false;
                TRENINGmetroTabControl6.Enabled = false;
                metroTile15.Enabled = false;
                metroTile17.Enabled = false;
                if (TRENINGmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренировок не найдено", "Тренировки не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrenerovka();
                }
                TRENINGtextBox3.Text = "";
            }
        }

        private void metroComboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
         
        }

        private void metroTile26_Click_1(object sender, EventArgs e)
        {
            ReportTrenerovka rep = new ReportTrenerovka();
            rep.Show();
        }

        private void metroTile35_Click(object sender, EventArgs e)
        {
            View_trenerovki view = new View_trenerovki();
            view.ShowDialog();
        }


     
        
      

        private void metroTile5_Click_1(object sender, EventArgs e)
        {
          
        }

        private void metroTile6_Click_2(object sender, EventArgs e)
        {
           
        }

        private void metroTile4_VisibleChanged(object sender, EventArgs e)
        {
            
        }

        public void pictureBox2_Click(object sender, EventArgs e)
        {
         // MOD_Trenerovka d = new MOD_Trenerovka();
            WindowState = FormWindowState.Normal;
            Height =-9999;
            Width = -9999;
        }

        private void metroTile13_Click_1(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();

            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT top 1 Тренировка.Идтренировка, тренировка.название, count(Тренировка.Идтренировка) as [В абонементах]
            FROM Тренировка INNER JOIN Абонемент ON Тренировка.Идтренировка = Абонемент.Идтренеровка
            GROUP BY Тренировка.Идтренировка, Тренировка.Название
            ORDER BY sum(Тренировка.Идтренировка)desc"), conn);
            TRENINGmetroTabControl3.Enabled = false;
            TRENINGmetroTabControl5.Enabled = false;
            TRENINGmetroTabControl4.Enabled = false;
            TRENINGmetroTabControl6.Enabled = false;
            TRENINGmetroTabControl7.Enabled = false;
            TRENINGmetroTabControl8.Enabled = false;
            adapter.Fill(ds);
            TRENINGmetroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            if (SPORTMmetroGrid2.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updTrenerovka();
            }


        }

        private void metroTile41_Click_1(object sender, EventArgs e)
        {
            MOD_Abonement ObjAbonementAdd = new MOD_Abonement();
            ObjAbonementAdd.textBox1.Text = "";
            ObjAbonementAdd.textBox2.Text = "";
            ObjAbonementAdd.textBox3.Text = "";
            ObjAbonementAdd.Text = "Добавить абонемент";
            ObjAbonementAdd.metroTile1.Text = "Добавить";
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Тренировка.Идтренировка, Тренировка.Название FROM Тренировка", con);
            ObjAbonementAdd.metroComboBox1.DisplayMember = "Название";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjAbonementAdd.metroComboBox1.DataSource = list.ToList();
            ObjAbonementAdd.metroComboBox1.DisplayMember = "Value";
            ObjAbonementAdd.metroComboBox1.ValueMember = "Key";
            con.Close();
            if (ObjAbonementAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    ABONmetroGrid1.Sort(ABONmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    idabonement = Convert.ToString(Convert.ToInt32(ABONmetroGrid1.Rows[ABONmetroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [абонемент]
                                                        ( Название, Цена, Количество_посещений, Идтренеровка)
                                                        VALUES(@st1,@st2,@st3,@st4)", con);
                    sss.Parameters.AddWithValue("st1", ObjAbonementAdd.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjAbonementAdd.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjAbonementAdd.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", Convert.ToInt32(ObjAbonementAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updAbonement();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile40_Click_1(object sender, EventArgs e)
        {
            MOD_Abonement ObjAbonementUpdate = new MOD_Abonement();
            con.Close();
            ObjAbonementUpdate.Text = "Редактировать абонемент";
            ObjAbonementUpdate.metroTile1.Text = "Редактировать";
            numbAbonement = Convert.ToInt32(ABONmetroGrid1.CurrentRow.Cells[0].Value);
            ObjAbonementUpdate.textBox1.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[1].Value);
            ObjAbonementUpdate.textBox2.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[2].Value);
            ObjAbonementUpdate.textBox3.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[3].Value);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Тренировка.Идтренировка, Тренировка.Название FROM Тренировка", con);
            ObjAbonementUpdate.metroComboBox1.DisplayMember = "Название";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjAbonementUpdate.metroComboBox1.DataSource = list.ToList();
            ObjAbonementUpdate.metroComboBox1.DisplayMember = "Value";
            ObjAbonementUpdate.metroComboBox1.ValueMember = "Key";
            con.Close();
            ObjAbonementUpdate.metroComboBox1.SelectedValue = ABONmetroGrid1.CurrentRow.Cells[5].Value;
            if (ObjAbonementUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    con.Close();
                    ABONmetroGrid1.Sort(ABONmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update абонемент set Название=@st1, Цена=@st2, Количество_посещений=@st3, Идтренеровка=@st4 where идабонемент=" + numbAbonement + "", con);
                    sss.Parameters.AddWithValue("st1", ObjAbonementUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjAbonementUpdate.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjAbonementUpdate.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", Convert.ToInt32(ObjAbonementUpdate.metroComboBox1.SelectedValue));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updAbonement();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }

        }

        private void metroTile39_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Close();
                con.Open();
                numbAbonement = Convert.ToInt32(ABONmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM абонемент 
                                                    WHERE идабонемент=" + numbAbonement + "", con);
                sss.ExecuteNonQuery();
                updAbonement();
                con.Close();
            }

        }

        private void metroTile38_Click_1(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updAbonement();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Абонемент";
            ExcelApp.Cells[1, 1] = "Название";
            ExcelApp.Cells[1, 2] = "Цена";
            ExcelApp.Cells[1, 3] = "Количество";
            ExcelApp.Cells[1, 4] = "Тренеровка";
            {
                for (int i = 1; i < ABONmetroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = ABONmetroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = ABONmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = ABONmetroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = ABONmetroGrid1.Rows[i - 1].Cells[4].Value;
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile37_Click_2(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trening.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, ABONmetroGrid1.RowCount + 1, 4);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Название";
                table.Cell(1, 2).Range.Text = "Цена"; ;
                table.Cell(1, 3).Range.Text = "Количество";
                table.Cell(1, 4).Range.Text = "Тренировка";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < ABONmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                }

                try
                {
                    doc.SaveAs(path);
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch { }
            }
            catch { }

        }

        private void metroTile36_Click_2(object sender, EventArgs e)
        {
            RepSportsmen resp = new RepSportsmen();
            resp.Show();
        }

        private void сброситьФильтрToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            updAbonement();
        }

        private void metroButton1_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox3.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE  Тренировка.Название='" + ABONtextBox3.Text + "'";
                sdaabonement = new OleDbDataAdapter(s, con);
                dtAbonement = new DataTable();
                sdaabonement.Fill(dtAbonement);
                ABONmetroGrid1.DataSource = dtAbonement;
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl6.Enabled = false;
                ABONmetroTabControl4.Enabled = false;
                metroTile19.Enabled = false;
                metroTile26.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких абонементов не найдено", "Абонемента не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updAbonement();
                }
                ABONtextBox3.Text = "";
            }
        }

        private void metroTile6_Click_3(object sender, EventArgs e)
        {

            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT top 1  абонемент.идабонемент,абонемент.название, count (абонемент.идабонемент) as [В продажах]
FROM Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент
 GROUP BY  абонемент.идабонемент,абонемент.название
            ORDER BY sum(абонемент.идабонемент)desc"), conn);
            adapter.Fill(ds);
            ABONmetroGrid1.DataSource = ds.Tables[0];
            
            conn.Close();
            ABONmetroTabControl4.Enabled = false;
            ABONmetroTabControl5.Enabled = false;
            ABONmetroTabControl6.Enabled = false;
            ABONmetroTabControl7.Enabled = false;
            ABONmetroTabControl8.Enabled = false;
            if (ABONmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updAbonement();
            }
        }

        private void metroTile28_Click_2(object sender, EventArgs e)
        {
            RepAbone_ent g = new RepAbone_ent();
            g.Show();
        }

        private void textBox8_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox8_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox3_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox8_KeyUp_1(object sender, KeyEventArgs e)
        {
            string s = @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                       WHERE[Абонемент.Название] LIKE '%" + ABONtextBox8.Text + "%'";
            sdaabonement = new OleDbDataAdapter(s, con);
            dtAbonement = new DataTable();
            sdaabonement.Fill(dtAbonement);
            ABONmetroGrid1.DataSource = dtAbonement;
            ABONmetroTabControl7.Enabled = false;
            ABONmetroTabControl5.Enabled = false;
            ABONmetroTabControl4.Enabled = false;
            metroTile27.Enabled = false;
            if (ABONmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемент не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ABONtextBox8.Text = "";
                updAbonement();
            }
        }

        private void metroButton4_Click_2(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            if (Convert.ToInt32(ABONmetroTextBox3.Text) < Convert.ToInt32(ABONmetroTextBox2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка  where Абонемент.Цена between {0} and {1}", ABONmetroTextBox3.Text, ABONmetroTextBox2.Text), conn);
                adapter.Fill(ds);
                ABONmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl5.Enabled = false;
                ABONmetroTabControl4.Enabled = false;
                metroTile35.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемента не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updAbonement();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальная цена не может быть больше конечной", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updAbonement();
            }
        }

        private void metroButton3_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE Абонемент.Название='" + ABONtextBox5.Text + "'";
                sdaabonement = new OleDbDataAdapter(s, con);
                dtAbonement = new DataTable();
                sdaabonement.Fill(dtAbonement);
                ABONmetroGrid1.DataSource = dtAbonement;
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl6.Enabled = false;
                ABONmetroTabControl4.Enabled = false;
                metroTile19.Enabled = false;
                metroTile5.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Абонемента не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updAbonement();
                }
                ABONtextBox5.Text = "";
            }
        }

        private void metroButton2_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE Абонемент.Количество_посещений=" +  ABONtextBox4.Text + "";
                sdaabonement = new OleDbDataAdapter(s, con);
                dtAbonement = new DataTable();
                sdaabonement.Fill(dtAbonement);
                ABONmetroGrid1.DataSource = dtAbonement;
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl6.Enabled = false;
                ABONmetroTabControl4.Enabled = false;
                metroTile26.Enabled = false;
                metroTile5.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких абонементов не найдено", "Абонемента не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updAbonement();
                }
                ABONtextBox4.Text = "";
            }
        }

        private void metroTile48_Click(object sender, EventArgs e)
        {
            MOD_SALE f = new MOD_SALE();
            f.ShowDialog();
        }

        private void metroTile51_Click(object sender, EventArgs e)
        {
            MOD_SALE ObjSaleAdd = new MOD_SALE();
            ObjSaleAdd.Text = "Добавить продажу";
            ObjSaleAdd.metroTile1.Text = "Добавить";

            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия FROM Сотрудник;", con);
            ObjSaleAdd.metroComboBox1.DisplayMember = "Фамилия";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjSaleAdd.metroComboBox1.DataSource = list.ToList();
            ObjSaleAdd.metroComboBox1.DisplayMember = "Value";
            ObjSaleAdd.metroComboBox1.ValueMember = "Key";
            con.Close();
            //////////////////////////////////////////////////
            con.Open();
            OleDbCommand cmd1 = new OleDbCommand("SELECT Спортсмен.Идспортсмен, Спортсмен.Фамилия FROM Спортсмен;", con);
            ObjSaleAdd.metroComboBox2.DisplayMember = "фамилия";
            OleDbDataReader reader1 = cmd1.ExecuteReader();
            Dictionary<int, string> list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int)reader1[0], (string)reader1[1]);
            }
            reader1.Close();
            cmd1.ExecuteNonQuery();
            ObjSaleAdd.metroComboBox2.DataSource = list1.ToList();
            ObjSaleAdd.metroComboBox2.DisplayMember = "Value";
            ObjSaleAdd.metroComboBox2.ValueMember = "Key";
            con.Close();
            ///////////////////////////////////////////////////////////////
            con.Open();
            OleDbCommand cmd2 = new OleDbCommand("SELECT Абонемент.Идабонемент, Абонемент.Название FROM Абонемент;", con);
            ObjSaleAdd.metroComboBox3.DisplayMember = "Название";
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            Dictionary<int, string> list2 = new Dictionary<int, string>();
            while (reader2.Read())
            {
                list2.Add((int)reader2[0], (string)reader2[1]);
            }
            reader2.Close();
            cmd2.ExecuteNonQuery();
            ObjSaleAdd.metroComboBox3.DataSource = list2.ToList();
            ObjSaleAdd.metroComboBox3.DisplayMember = "Value";
            ObjSaleAdd.metroComboBox3.ValueMember = "Key";
            con.Close();


            if (ObjSaleAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    SALEmetroGrid1.Sort(SALEmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    idSALE = Convert.ToString(Convert.ToInt32(SALEmetroGrid1.Rows[SALEmetroGrid1.RowCount - 1].Cells[0].Value) + 1);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand(@"INSERT INTO [Продажа_абонемента]
                                                        ( Идсотрудник, Идспортсмен, Идабонемент, Дата_начала,Дата_окончания)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5)", con);
                    sss.Parameters.AddWithValue("st1", Convert.ToInt32(ObjSaleAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjSaleAdd.metroComboBox2.SelectedValue.ToString()));
                    sss.Parameters.AddWithValue("st3", Convert.ToInt32(ObjSaleAdd.metroComboBox3.SelectedValue.ToString()));
                    sss.Parameters.AddWithValue("st4", Convert.ToDateTime(ObjSaleAdd.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st5", Convert.ToDateTime(ObjSaleAdd.metroDateTime2.Text));

                    sss.ExecuteNonQuery();
                    con.Close();
                    updSale();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile50_Click(object sender, EventArgs e)
        {
            MOD_SALE ObjSaleUpdate = new MOD_SALE();
            con.Close();
            ObjSaleUpdate.Text = "Редактировать продажи";
            ObjSaleUpdate.metroTile1.Text = "Редактировать";
            numbstrSALE = Convert.ToInt32(SALEmetroGrid1.CurrentRow.Cells[0].Value);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия FROM Сотрудник;", con);
            ObjSaleUpdate.metroComboBox1.DisplayMember = "Фамилия";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, string> list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (string)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjSaleUpdate.metroComboBox1.DataSource = list.ToList();
            ObjSaleUpdate.metroComboBox1.DisplayMember = "Value";
            ObjSaleUpdate.metroComboBox1.ValueMember = "Key";
            con.Close();

            ObjSaleUpdate.metroComboBox1.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[7].Value;

            //////////////////////////////////////////////////
            con.Open();
            OleDbCommand cmd1 = new OleDbCommand("SELECT Спортсмен.Идспортсмен, Спортсмен.Фамилия FROM Спортсмен;", con);
            ObjSaleUpdate.metroComboBox2.DisplayMember = "фамилия";
            OleDbDataReader reader1 = cmd1.ExecuteReader();
            Dictionary<int, string> list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int)reader1[0], (string)reader1[1]);
            }
            reader1.Close();
            cmd1.ExecuteNonQuery();
            ObjSaleUpdate.metroComboBox2.DataSource = list1.ToList();
            ObjSaleUpdate.metroComboBox2.DisplayMember = "Value";
            ObjSaleUpdate.metroComboBox2.ValueMember = "Key";
            con.Close();
            ObjSaleUpdate.metroComboBox2.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[8].Value;

            ///////////////////////////////////////////////////////////////
            con.Open();
            OleDbCommand cmd2 = new OleDbCommand("SELECT Абонемент.Идабонемент, Абонемент.Название FROM Абонемент;", con);
            ObjSaleUpdate.metroComboBox3.DisplayMember = "Название";
            OleDbDataReader reader2 = cmd2.ExecuteReader();
            Dictionary<int, string> list2 = new Dictionary<int, string>();
            while (reader2.Read())
            {
                list2.Add((int)reader2[0], (string)reader2[1]);
            }
            reader2.Close();
            cmd2.ExecuteNonQuery();
            ObjSaleUpdate.metroComboBox3.DataSource = list2.ToList();
            ObjSaleUpdate.metroComboBox3.DisplayMember = "Value";
            ObjSaleUpdate.metroComboBox3.ValueMember = "Key";
            con.Close();
            ObjSaleUpdate.metroComboBox3.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[9].Value;

            ObjSaleUpdate.metroDateTime1.Text = Convert.ToString(SALEmetroGrid1.CurrentRow.Cells[4].Value);
            ObjSaleUpdate.metroDateTime2.Text = Convert.ToString(SALEmetroGrid1.CurrentRow.Cells[5].Value);
          
            if (ObjSaleUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    con.Close();
                    SALEmetroGrid1.Sort(SALEmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update Продажа_абонемента set Идсотрудник=@st1, Идспортсмен=@st2, Идабонемент=@st3, Дата_начала=@st4,Дата_окончания=@st5 where идпродажа=" + numbstrSALE + "", con);
                    sss.Parameters.AddWithValue("st1", Convert.ToInt32(ObjSaleUpdate.metroComboBox1.SelectedValue));
                    sss.Parameters.AddWithValue("st2", Convert.ToInt32(ObjSaleUpdate.metroComboBox2.SelectedValue));
                    sss.Parameters.AddWithValue("st3", Convert.ToInt32(ObjSaleUpdate.metroComboBox3.SelectedValue));
                    sss.Parameters.AddWithValue("st4", Convert.ToDateTime(ObjSaleUpdate.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st5", Convert.ToDateTime(ObjSaleUpdate.metroDateTime2.Text));
                    sss.ExecuteNonQuery();
                    con.Close();
                    updSale();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this, ex.Message, "Ошибка");
                }
        }

        private void metroTile49_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Close();
                con.Open();
                numbstrSALE = Convert.ToInt32(SALEmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM продажа_абонемента 
                                                    WHERE идпродажа=" + numbstrSALE + "", con);
                sss.ExecuteNonQuery();
                updSale();
                con.Close();
            }
        }

        private void metroTile47_Click(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.updSale();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Продажи";
            ExcelApp.Cells[1, 1] = "Сотрудник";
            ExcelApp.Cells[1, 2] = "Абонемент";
            ExcelApp.Cells[1, 3] = "Количество_посещений";
            ExcelApp.Cells[1, 4] = "Дата_начала";
            ExcelApp.Cells[1, 5] = "Дата_окончания";
            ExcelApp.Cells[1, 6] = "Фамилия";
        
            {
                for (int i = 1; i < SALEmetroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = SALEmetroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = SALEmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = SALEmetroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = SALEmetroGrid1.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = SALEmetroGrid1.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = SALEmetroGrid1.Rows[i - 1].Cells[6].Value;
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile46_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/SALE.docx";
            var Wordapp = new Microsoft.Office.Interop.Word.Application();

            Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, SALEmetroGrid1.RowCount + 1, 6);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Сотрудник";
                table.Cell(1, 2).Range.Text = "Абонемент"; ;
                table.Cell(1, 3).Range.Text = "Количество_посещений";
                table.Cell(1, 4).Range.Text = "Дата_начала";
                table.Cell(1, 5).Range.Text = "Дата_окончания";
                table.Cell(1, 6).Range.Text = "Фамилия";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < SALEmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = EMPLmetroGrid1.Rows[i - 1].Cells[6].Value.ToString();
                }

                try
                {
                    doc.SaveAs(path);
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                catch { }
            }
            catch { }
        }

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char)Keys.Back)

                {
                    e.Handled = true;
                    DialogResult result = MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_3(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_3(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
            ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void textBox6_KeyUp_1(object sender, KeyEventArgs e)
        {

            string s = @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник 
                       WHERE[Абонемент.Название] LIKE '%" + SALEtextBox6.Text + "%'";
            sdaSALE = new OleDbDataAdapter(s, con);
            dtSALE = new DataTable();
            sdaSALE.Fill(dtSALE);
            SALEmetroGrid1.DataSource = dtSALE;
            SALEmetroTabControl4.Enabled = false;
            SALEmetroTabControl5.Enabled = false;
            SALEmetroTabControl7.Enabled = false;
            metroTile42.Enabled = false;
            metroTile43.Enabled = false;
            if (SALEmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемента не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                SALEtextBox6.Text = "";
                updSale();
            }
        }

        private void metroButton5_Click_2(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            //  conn.Close();
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = SALEmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = SALEmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(SALEmetroDateTime2.Text) < Convert.ToDateTime(SALEmetroDateTime1.Text))
            {
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl5.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                metroTile42.Enabled = false;
                metroTile44.Enabled = false;
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
WHERE Продажа_абонемента.Дата_начала Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                SALEmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();

                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Даты не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSale();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updEmployee();
            }
        }

        private void metroButton4_Click_3(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            if (Convert.ToInt32(SALEmetroTextBox3.Text) < Convert.ToInt32(SALEmetroTextBox2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник where Абонемент.Количество_посещений between {0} and {1}", SALEmetroTextBox3.Text, SALEmetroTextBox2.Text), conn);
                adapter.Fill(ds);
                SALEmetroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl5.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                metroTile43.Enabled = false;
                metroTile44.Enabled = false;

                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Количество не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSale();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальное количество не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updSale();
            }
        }

        private void metroButton3_Click_4(object sender, EventArgs e)
        {
            if (SALEtextBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl6.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                metroTile40.Enabled = false;
               
                string s = @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
                           WHERE Спортсмен.Фамилия='" + SALEtextBox5.Text + "'";
                sdaSALE = new OleDbDataAdapter(s, con);
                dtSALE = new DataTable();
                sdaSALE.Fill(dtSALE);
                SALEmetroGrid1.DataSource = dtSALE;

                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSale();
                }
                SALEtextBox5.Text = "";
            }
        }

        private void metroButton2_Click_4(object sender, EventArgs e)
        {
            if (SALEtextBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl6.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                metroTile41.Enabled = false;
                string s = @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
                           WHERE Сотрудник.Фамилия='" + SALEtextBox4.Text + "'";
                sdaSALE = new OleDbDataAdapter(s, con);
                dtSALE = new DataTable();
                sdaSALE.Fill(dtSALE);
                SALEmetroGrid1.DataSource = dtSALE;

                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updSale();
                }
                SALEtextBox4.Text = "";
            }
        }

        private void metroTile38_Click_2(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Абонемент.Название AS Абонемент, Сотрудник.Фамилия  & Сотрудник.Имя AS [Реализовавший сотрудник], Спортсмен.Фамилия  & Спортсмен.Имя AS Спортсмен, Продажа_абонемента.Дата_начала as Начало,  Продажа_абонемента.Дата_окончания as Конец
                                                                          FROM Спортсмен 
                                                                          INNER JOIN (Сотрудник 
                                                                          INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                          ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник) 
                                                                          ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;"), conn);
            adapter.Fill(ds);
            SALEmetroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            //EnableOFF();
            SALEmetroTabControl4.Enabled = false;
            SALEmetroTabControl5.Enabled = false;
            SALEmetroTabControl6.Enabled = false;
            SALEmetroTabControl9.Enabled = false;
            SALEmetroTabControl7.Enabled = false;
            if (SALEmetroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                updSale();
            }
        }

        private void сброситьФильтрToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            updSale();
        }

        private void metroTile45_Click(object sender, EventArgs e)
        {
            RepSportsmen resp = new RepSportsmen();
            resp.Show();
        }

        private void metroTile28_Click_3(object sender, EventArgs e)
        {
            try
            {
                VISITS f = new VISITS();
                f.ShowDialog();
            }
            catch { }
        }
    }
}







