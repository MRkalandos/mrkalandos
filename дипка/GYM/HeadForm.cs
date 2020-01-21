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
        }
        public int numbstrokasportsmen = 0;
        public int numbstrokemployee = 0;
        public int numbstroktrener = 0;
        public string idEmployee;
        public string idsportsmen;
        public string idtrener;
        public DataTable dtEmployee;
        public DataTable dtSportsmen;
        public DataTable dtTrener;
        string filename;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
       public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
        public OleDbDataAdapter sdaEmployee;
        public OleDbDataAdapter sdasportsmen;
        public OleDbDataAdapter sdatrener;

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

        public void updTrener()
        {
            try
            {
                sdatrener = new OleDbDataAdapter(@"SELECT Тренер.*
                                                    FROM Тренер;", con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                metroGrid1.DataSource = dtTrener;
                metroGrid1.Columns[0].Visible = false;
                metroGrid1.Columns[8].Visible = false;
                textBox8.Text = "";
                textBox5.Text = "";
                textBox4.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                metroTextBox2.Text = "";
                metroTextBox3.Text = "";
                metroGrid1.Select();
                metroGrid1.AllowUserToAddRows = false;
                metroTile16.Visible = true;
                pictureBox2.Visible = true;
                metroTile11.Enabled = true;
                metroTile12.Enabled = true;
                metroTile10.Enabled = true;
                metroTabControl4.Enabled = true;
                metroTabControl3.Enabled = true;
                metroTabControl6.Enabled = true;
                metroTile7.Enabled = true;
                metroTile8.Enabled = true;
                metroTile9.Enabled = true;
                metroTabControl5.Enabled = true;
                metroTabControl8.Enabled = true;
                
                metroTabControl5.Enabled = true;


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


            ///////////////////////////////////////////////trener///////////////////////////////////
            OleDbConnection conTrener = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sdatrener = new OleDbDataAdapter(@"SELECT тренер.*
                                                 FROM тренер;", conSportsmen);
            dtTrener = new DataTable();
            sdatrener.Fill(dtTrener);
            metroGrid1.DataSource = dtTrener;

            metroGrid1.Columns[0].Visible = false;
            SPORTMmetroGrid2.Columns[0].Visible = false;
            EMPLmetroGrid1.Columns[0].Visible = false;
        }

        private void HeadeForm_Load(object sender, EventArgs e)
        {

            SPORTMmetroComboBox1.SelectedIndex = 0;
            
            try
            {
                updTrener();
                updEmployee();
                updSportsmen();
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
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите закрыть программу?", "Подтверждение выхода", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Environment.Exit(0);
                //int hWnd = FindWindow("Shell_TrayWnd", "");
               // ShowWindow(hWnd, SW_SHOW); //показываем панель задач
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
            string n = Microsoft.VisualBasic.Interaction. InputBox("Введите Фамилию спортсмена: ", "Запрос",
                "", 500, 500);
            OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@" SELECT Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество, Учет_посещений.Дата,Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, 
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
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    idtrener = Convert.ToString(Convert.ToInt32(metroGrid1.Rows[metroGrid1.RowCount - 1].Cells[0].Value) + 1);
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
            numbstroktrener = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
            ObjTrenerUpdate.textBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[1].Value);
            ObjTrenerUpdate.textBox2.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[2].Value);
            ObjTrenerUpdate.textBox3.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[3].Value);
            ObjTrenerUpdate.metroTextBox4.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[4].Value);
            ObjTrenerUpdate.maskedTextBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[5].Value);
            ObjTrenerUpdate.metroDateTime1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[6].Value);
            ObjTrenerUpdate.metroTextBox5.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[7].Value);
            ObjTrenerUpdate.metroTextBox6.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[8].Value);
            ObjTrenerUpdate.metroTextBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[9].Value);
            if (ObjTrenerUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
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
                numbstroktrener = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
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
                for (int i = 1; i < metroGrid1.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = metroGrid1.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = metroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = metroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = metroGrid1.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = metroGrid1.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = metroGrid1.Rows[i - 1].Cells[6].Value;
                    ExcelApp.Cells[i + 1, 7] = metroGrid1.Rows[i - 1].Cells[7].Value;
                    ExcelApp.Cells[i + 1, 8] = metroGrid1.Rows[i - 1].Cells[9].Value;
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
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, metroGrid1.RowCount + 1, 8);
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
                for (int i = 1; i < metroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = metroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = metroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = metroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = metroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = metroGrid1.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = metroGrid1.Rows[i - 1].Cells[6].Value.ToString();
                    table.Cell(i + 1, 7).Range.Text = metroGrid1.Rows[i - 1].Cells[7].Value.ToString();
                    table.Cell(i + 1, 8).Range.Text = metroGrid1.Rows[i - 1].Cells[9].Value.ToString();
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
                       WHERE[Фамилия] LIKE '%" + textBox8.Text + "%'";
            sdatrener = new OleDbDataAdapter(s, con);
            dtTrener = new DataTable();
            sdatrener.Fill(dtTrener);
            metroGrid1.DataSource = dtTrener;
            metroTile11.Enabled=false;
            metroTile10.Enabled = false;
            metroTabControl4.Enabled = false;
            metroTabControl3.Enabled = false;
            metroTabControl6.Enabled = false;
            if (metroGrid1.RowCount == 0)
            {
                MetroFramework.MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox8.Text = "";
                updTrener();
            }
        }

        private void metroButton5_Click_1(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = metroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = metroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(metroDateTime2.Text) < Convert.ToDateTime(metroDateTime1.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter
                (String.Format(@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                                WHERE Дата_рождения Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                metroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTile12.Enabled = false;
                metroTile10.Enabled = false;
                metroTabControl4.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl6.Enabled = false;

                if (metroGrid1.RowCount == 0)
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
            if (Convert.ToInt32(metroTextBox3.Text) < Convert.ToInt32(metroTextBox2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                               where оклад between {0} and {1}", metroTextBox3.Text, metroTextBox2.Text), conn);
                adapter.Fill(ds);
                metroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTile12.Enabled = false;
                metroTile11.Enabled = false;
                metroTabControl4.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl6.Enabled = false;
                if (metroGrid1.RowCount == 0)
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
            if (textBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Фамилия='" + textBox5.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                metroGrid1.DataSource = dtTrener;
                metroGrid1.Columns[8].Visible = false;
                //  metroGrid1.Columns[0].Visible = false;
                metroTile7.Enabled = false;
                metroTile8.Enabled = false;
                metroTabControl5.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl6.Enabled = false;
                if (metroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                textBox5.Text = "";
            }
        }

        private void metroButton2_Click_1(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Имя='" + textBox4.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                metroGrid1.DataSource = dtTrener;
                metroGrid1.Columns[8].Visible = false;
               // metroGrid1.Columns[0].Visible = false;
                metroTile7.Enabled = false;
                metroTile9.Enabled = false;
                metroTabControl5.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl6.Enabled = false;

                if (metroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                textBox4.Text = "";
            }
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
           // EMPLmetroTile3.Visible = true;
            if ((textBox3.Text == "") || (textBox2.Text == ""))
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string s = @"SELECT Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                               WHERE Фамилия='" + textBox3.Text + "' and Имя='" + textBox2.Text + "'";
                sdatrener = new OleDbDataAdapter(s, con);
                dtTrener = new DataTable();
                sdatrener.Fill(dtTrener);
                metroGrid1.DataSource = dtTrener;
                metroGrid1.Columns[8].Visible = false;
                metroTile8.Enabled = false;
                metroTile9.Enabled = false;
                metroTabControl5.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl6.Enabled = false;
                if (metroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updTrener();
                }
                textBox3.Text = "";
                textBox2.Text = "";
            }
        }

        private void metroGrid1_SelectionChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (metroGrid1.CurrentRow != null)
                {
                    pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                    textBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[8].Value);
                    pictureBox2.Load(Application.StartupPath + @"\PhotoTrener\" + textBox1.Text);
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
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Тренер.Фамилия, Тренер.Имя, Тренер.Отчество, Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания,тренер.фото
FROM Тренер INNER JOIN
(Спортсмен INNER JOIN 
(Абонемент INNER JOIN Продажа_абонемента
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент)
ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Тренер.Идтренер = Абонемент.Идтренер;"), conn);

                metroTile16.Visible = false;
                pictureBox2.Visible=false;

                adapter.Fill(ds);
                metroGrid1.DataSource = ds.Tables[0];

                metroTabControl8.Enabled = false;
                metroTabControl6.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTabControl4.Enabled = false;
                metroTabControl5.Enabled = false;
                conn.Close();


                if (metroGrid1.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    updTrener();
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
        }
    }
}







