using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
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
        public string idEmployee;
        public string idsportsmen;
        public DataTable dtEmployee;
        public DataTable dtSportsmen;
        string filename;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        public OleDbDataAdapter sdaEmployee;
        public OleDbDataAdapter sdasportsmen;

        public void upd()
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
                sdasportsmen = new OleDbDataAdapter(@"SELECT Спортсмен.*
                                                 FROM Спортсмен;", con);
                dtSportsmen = new DataTable();
                sdasportsmen.Fill(dtSportsmen);
                metroGrid2.DataSource = dtSportsmen;
                metroGrid2.Columns[0].Visible = false;
                metroGrid2.Select();
                metroGrid2.AllowUserToAddRows = false;

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
            metroGrid2.DataSource = dtSportsmen;

        }

        private void HeadeForm_Load(object sender, EventArgs e)
        {
            this.EMPLзарплата_сотрудникаTableAdapter.Fill(this.EMPLdS_Money.Зарплата_сотрудника);
            try
            {
                upd();
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
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
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
            
                upd();
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
            ObjEmployeeAdd.metroTextBox6.Text = "";
            ObjEmployeeAdd.metroTextBox5.Text = "";
            ObjEmployeeAdd.Text = "Добавить сотрудника";
            ObjEmployeeAdd.metroTile1.Text = "Добавить";
            ObjEmployeeAdd.metroComboBox1.DataSource = EMPLbindingSource1;
            ObjEmployeeAdd.metroComboBox1.ValueMember = "Идзарплата";
            ObjEmployeeAdd.metroComboBox1.DisplayMember ="Зарплата";
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
                    sss.Parameters.AddWithValue("st7", Convert.ToInt16(ObjEmployeeAdd.metroTextBox5.Text));
                    sss.Parameters.AddWithValue("st8", ObjEmployeeAdd.metroTextBox6.Text);
                    sss.Parameters.AddWithValue("st9", Convert.ToInt16(ObjEmployeeAdd.metroComboBox1.SelectedValue.ToString()));
                    sss.ExecuteNonQuery();
                    con.Close();
                    upd();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this,ex.Message ,"Ошибка");
                }
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            MOD_Employee ObjEmployeeUpdate = new MOD_Employee();

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
            OleDbConnection con2 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            con2.Open();
            OleDbCommand cmd = new OleDbCommand("select идзарплата,зарплата from зарплата_сотрудника", con2);
            ObjEmployeeUpdate.metroComboBox1.DisplayMember = "Зарплата";
            OleDbDataReader reader = cmd.ExecuteReader();
            Dictionary<int, int> list = new Dictionary<int, int>();
            while (reader.Read())
            {
                list.Add((int)reader[0], (int)reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            ObjEmployeeUpdate.metroComboBox1.DataSource = list.ToList();
            ObjEmployeeUpdate.metroComboBox1.DisplayMember = "Value";
            ObjEmployeeUpdate.metroComboBox1.ValueMember = "Key";
            ObjEmployeeUpdate.metroComboBox1.SelectedValue = EMPLmetroGrid1.CurrentRow.Cells[10].Value;
            if (ObjEmployeeUpdate.ShowDialog() == DialogResult.OK)
                try
                {
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
                    upd();
                }
                catch (Exception ex)
                {
                    MetroFramework.MetroMessageBox.Show(this,ex.Message,"Ошибка");
                }
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                con.Open();
                numbstrokemployee = Convert.ToInt32(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM сотрудник 
                                                    WHERE идсотрудник=" + numbstrokemployee + "", con);
                sss.ExecuteNonQuery();
                upd();
                con.Close();
            }
        }

        private void metroTile8_Click(object sender, EventArgs e)
        {

            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
           this.upd();
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
                for (int i = 1; i < EMPLmetroGrid1.Rows.Count+1; i++)
                {
                     ExcelApp.Cells[i + 1, 1] = EMPLmetroGrid1.Rows[i-1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = EMPLmetroGrid1.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = EMPLmetroGrid1.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = EMPLmetroGrid1.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = EMPLmetroGrid1.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1,6] = EMPLmetroGrid1.Rows[i - 1].Cells[6].Value;
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
           // Wordapp.Visible = true;
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
                doc.SaveAs(path);
               
                try
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Wordapp.Visible = true;//открывает прогу
                    //doc.Close();
                  //  Wordapp.Quit();
                }
                catch { }
            }
            catch { }
        }

        private void metroTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
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
                upd();
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
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            OleDbConnection conn = new OleDbConnection(conString);
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
                    upd();
                }
        }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                upd();
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
                    upd();         
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
                    upd();
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
                    upd();
                }
                EMPLtextBox5.Text = "";
                EMPLtextBox8.Text = "";
            }
        }

        private void сброситьФильтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upd();
        }

        private void metroGrid1_SelectionChanged(object sender, EventArgs e)
        {
            if (EMPLmetroGrid1.CurrentRow != null)
            {
               EMPLpictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                metroTextBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[8].Value);
                EMPLpictureBox1.Load(Application.StartupPath + @"\model\" + metroTextBox1.Text);
            }
        }

        private void metroTile12_Click(object sender, EventArgs e)
        {
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
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

                upd();
            }
        }

        private void metroTile17_Click(object sender, EventArgs e)
        {
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
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
                upd();
            }
        }

        private void metroTile19_Click(object sender, EventArgs e)
        {
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
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

                upd();
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
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
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
                    upd();
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "\nНачальная зарплата не может быть больше конечной", "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                upd();
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
                    metroGrid2.Sort(metroGrid2.Columns[1], ListSortDirection.Ascending);
                    idsportsmen = Convert.ToString(Convert.ToInt32(metroGrid2.Rows[metroGrid2.RowCount - 1].Cells[0].Value) + 1);
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
                    upd();
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
            numbstrokasportsmen = Convert.ToInt32(metroGrid2.CurrentRow.Cells[0].Value);
            ObjSportsmenUpdate.textBox1.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[1].Value);
            ObjSportsmenUpdate.textBox2.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[2].Value);
            ObjSportsmenUpdate.textBox3.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[3].Value);
            ObjSportsmenUpdate.maskedTextBox1.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[4].Value);
            ObjSportsmenUpdate.metroDateTime1.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[5].Value);
            ObjSportsmenUpdate.metroComboBox1.Text = Convert.ToString(metroGrid2.CurrentRow.Cells[6].Value);
            if (ObjSportsmenUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    metroGrid2.Sort(metroGrid2.Columns[1], ListSortDirection.Ascending);
                    con.Open(); 
                    OleDbCommand sss  = new OleDbCommand("update спортсмен set Фамилия=@st1, Имя=@st2, Отчество=@st3, Телефон=@st4, Дата_рождения=@st5,Пол=@st6 where идспортсмен=" + numbstrokasportsmen + "", con);
                    sss.Parameters.AddWithValue("st1", ObjSportsmenUpdate.textBox1.Text);
                    sss.Parameters.AddWithValue("st2", ObjSportsmenUpdate.textBox2.Text);
                    sss.Parameters.AddWithValue("st3", ObjSportsmenUpdate.textBox3.Text);
                    sss.Parameters.AddWithValue("st4", ObjSportsmenUpdate.maskedTextBox1.Text);
                    sss.Parameters.AddWithValue("st5", Convert.ToDateTime(ObjSportsmenUpdate.metroDateTime1.Text));
                    sss.Parameters.AddWithValue("st6", ObjSportsmenUpdate.metroComboBox1.Text);
                    sss.ExecuteNonQuery();
                    con.Close();
                    upd();
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
                numbstrokasportsmen = Convert.ToInt32(metroGrid2.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM спортсмен 
                                                    WHERE идспортсмен=" + numbstrokasportsmen + "", con);
                sss.ExecuteNonQuery();
                upd();
                con.Close();
            }
        }

        private void metroTile37_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.upd();
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
                for (int i = 1; i < metroGrid2.Rows.Count + 1; i++)
                {
                    ExcelApp.Cells[i + 1, 1] = metroGrid2.Rows[i - 1].Cells[1].Value;
                    ExcelApp.Cells[i + 1, 2] = metroGrid2.Rows[i - 1].Cells[2].Value;
                    ExcelApp.Cells[i + 1, 3] = metroGrid2.Rows[i - 1].Cells[3].Value;
                    ExcelApp.Cells[i + 1, 4] = metroGrid2.Rows[i - 1].Cells[4].Value;
                    ExcelApp.Cells[i + 1, 5] = metroGrid2.Rows[i - 1].Cells[5].Value;
                    ExcelApp.Cells[i + 1, 6] = metroGrid2.Rows[i - 1].Cells[6].Value;

                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile36_Click(object sender, EventArgs e)
        {
            MetroFramework.MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Sportsmen.docx";
                        var Wordapp = new Microsoft.Office.Interop.Word.Application();
            // Wordapp.Visible = true;
            Microsoft.Office.Interop.Word.Document doc = Wordapp.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, metroGrid2.RowCount + 1, 6);
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
                for (int i = 1; i < metroGrid2.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = metroGrid2.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = metroGrid2.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = metroGrid2.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = metroGrid2.Rows[i - 1].Cells[4].Value.ToString();
                    table.Cell(i + 1, 5).Range.Text = metroGrid2.Rows[i - 1].Cells[5].Value.ToString();
                    table.Cell(i + 1, 6).Range.Text = metroGrid2.Rows[i - 1].Cells[6].Value.ToString();
                                   }
                doc.SaveAs(path);

                try
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Wordapp.Visible = true;//открывает прогу
                                           //doc.Close();
                                           //  Wordapp.Quit();
                }
                catch { }
            }
            catch { }
        }
    }
}







