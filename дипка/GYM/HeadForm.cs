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
        public int rez = 0;
        public string id;
        public DataTable dt;
        string filename;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        public OleDbDataAdapter sda;

        public void upd()
        {
            sda = new OleDbDataAdapter(@"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата, сотрудник.идзарплата  
                                        FROM Зарплата_сотрудника 
                                        INNER JOIN Сотрудник 
                                        ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", con);
            dt = new DataTable();
            sda.Fill(dt);
            metroGrid1.DataSource = dt;
            metroTile3.Visible = true;
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox8.Text = "";
            metroTextBox6.Text = "";
            metroTextBox7.Text = "";
            metroTabControl3.Enabled = true;
            metroTabControl4.Enabled = true;
            metroTabControl5.Enabled = true;
            metroTabControl6.Enabled = true;
            metroTabControl5.Enabled = true;
            metroTabControl7.Enabled = true;
            metroTabControl8.Enabled = true;
            metroTile10.Enabled = true;
            metroTile16.Enabled = true;
            metroTabControl8.Enabled = true;
            metroTabControl7.Enabled = true;
            metroTabControl8.Enabled = true;
            metroTile11.Enabled = true;
            metroTile16.Enabled = true;
            metroTile14.Enabled = true;
            metroTile15.Enabled = true;
            metroGrid1.Columns[8].Visible = false;
            metroGrid1.Columns[0].Visible = false;
            metroGrid1.Columns[10].Visible = false;
            metroGrid1.Select();
            metroGrid1.AllowUserToAddRows = false;
        }

        public void updfilename()
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filename + "");
            sda = new OleDbDataAdapter(@"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата
                                       FROM Зарплата_сотрудника 
                                       INNER JOIN Сотрудник   
                                       ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", con);
            dt = new DataTable();
            sda.Fill(dt);
            metroGrid1.DataSource = dt;
        }

        private void HeadeForm_Load(object sender, EventArgs e)
        {

            // TODO: данная строка кода позволяет загрузить данные в таблицу "dS_Money.Зарплата_сотрудника". При необходимости она может быть перемещена или удалена.
            this.зарплата_сотрудникаTableAdapter.Fill(this.dS_Money.Зарплата_сотрудника);
            try
            {
                upd();
            }
            catch
            {
                MessageBox.Show("Файл базы данных не найден укажите новый путь", "Файл БД не найден");
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
                    MessageBox.Show("Вы не выбрали файл БД, поэтому выбранно последнее выбранное место БД", "Не указана БД", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                }
            }
        }

        private void metroPanel1_Paint(object sender, PaintEventArgs e)
        {
        }

        private void metroPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void metroLabel4_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel7_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel8_Click(object sender, EventArgs e)
        {

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
            metroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            metroGrid1.Columns[0].Visible = false;
            metroGrid1.Columns[8].Visible = false;
            metroTabControl3.Enabled = false;
            if (metroGrid1.RowCount == 0)
            {
                MessageBox.Show("Запрос не дал результатов", "Запрос пуст");
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

            ObjEmployeeAdd.metroComboBox1.DataSource = bindingSource1;
            ObjEmployeeAdd.metroComboBox1.ValueMember = "Идзарплата";
            ObjEmployeeAdd.metroComboBox1.DisplayMember ="Зарплата";
            if (ObjEmployeeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    id = Convert.ToString(Convert.ToInt32(metroGrid1.Rows[metroGrid1.RowCount - 1].Cells[0].Value) + 1);
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
                    MessageBox.Show(ex.Message);
                }
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            MOD_Employee ObjEmployeeUpdate = new MOD_Employee();

            ObjEmployeeUpdate.Text = "Редактировать сотрудника";
            ObjEmployeeUpdate.metroTile1.Text = "Редактировать";
            rez = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
            ObjEmployeeUpdate.textBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[1].Value);
            ObjEmployeeUpdate.textBox2.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[2].Value);
            ObjEmployeeUpdate.textBox3.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[3].Value);
            ObjEmployeeUpdate.metroTextBox4.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[4].Value);
            ObjEmployeeUpdate.maskedTextBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[5].Value);
            ObjEmployeeUpdate.metroDateTime1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[6].Value);
            ObjEmployeeUpdate.metroTextBox5.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[7].Value);
            ObjEmployeeUpdate.metroTextBox6.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[8].Value);
            //  ObjEmployeeUpdate.comboBox1.ValueMember = Convert.ToString( dt);
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
            ObjEmployeeUpdate.metroComboBox1.SelectedValue = metroGrid1.CurrentRow.Cells[10].Value;
            if (ObjEmployeeUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    metroGrid1.Sort(metroGrid1.Columns[1], ListSortDirection.Ascending);
                    con.Open();
                    OleDbCommand sss = new OleDbCommand("update Сотрудник set Фамилия=@st1, Имя=@st2, Отчество=@st3, Должность=@st4,Телефон=@st5, Дата=@st6,Пароль=@st7,Фото=@st8,Идзарплата=@st9 where идсотрудник=" + rez + "", con);
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
                    MessageBox.Show(ex.Message);
                }
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Вы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                con.Open();
                rez = Convert.ToInt32(metroGrid1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM сотрудник 
                                                    WHERE идсотрудник=" + rez + "", con);
                sss.ExecuteNonQuery();
                upd();
                con.Close();
            }
        }

        private void metroTile8_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(delegate (object obj)
            {
                MessageBox.Show("Выполняется формирование отчета, это займёт около 10 секунд");
            }).Start();
            this.upd();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 15;
            ExcelWorkSheet.Name = "Сотрудники";
            ExcelApp.Cells[1, 1] = "IDСотрудника";
            ExcelApp.Cells[1, 2] = "Фамилия";
            ExcelApp.Cells[1, 3] = "Имя";
            ExcelApp.Cells[1, 4] = "Отчество";
            ExcelApp.Cells[1, 5] = "Должность";
            ExcelApp.Cells[1, 6] = "Телефон";
            ExcelApp.Cells[1, 7] = "Дата";
            ExcelApp.Cells[1, 8] = "Пароль";
            ExcelApp.Cells[1, 9] = "Фото";
            ExcelApp.Cells[1, 10] = "Зарплата";
            {
                for (int i = 0; i < metroGrid1.Rows.Count; i++)
                {
                    for (int j = 0; j < metroGrid1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = metroGrid1.Rows[i].Cells[j].Value;
                    }
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
        }

        private void metroTile9_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(delegate (object obj)
            {
                MessageBox.Show("Выполняется формирование отчета, это займёт около 10 секунд");
            }).Start();
            //Переключение активности кнопки
            var setButtonEnabled = new Action<bool>(enabled => metroTile9.Enabled = enabled);
            //Дективация кнопки
            setButtonEnabled(false);
            //По завершении экспорта вызываем переключение активности кнопки
            WordExporter.ExportCompleted += () =>
            {
                //Поскольку это действие выполняется в другом потоке, то используем InvokeRequired
                if (InvokeRequired)
                    BeginInvoke(setButtonEnabled, true);
                else
                    setButtonEnabled(true);
            };
            //Начинаем экспорт
            WordExporter.Export(metroGrid1, Path.GetFullPath("exported.doc"));
        }

        private void metroTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            metroTabControl5.Enabled = false;
           metroTile11.Enabled = false;
            metroTabControl7.Enabled = false;
            metroTabControl8.Enabled = false;
            metroTile16.Enabled = false;
            metroTile3.Visible = true;
            string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                       FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                       WHERE[Фамилия] LIKE '%" + textBox2.Text + "%'";
            sda = new OleDbDataAdapter(s, con);
            dt = new DataTable();
            sda.Fill(dt);
            metroGrid1.DataSource = dt;

            if (metroGrid1.RowCount == 0)
            {
                MessageBox.Show("Запись не найдена","Сотрудника не найдено");
                textBox2.Text = "";
                upd();
            }
        }

        public void EnableOFF()
        {
            metroTabControl3.Enabled = false;
            metroTabControl4.Enabled = false;
            metroTabControl5.Enabled = false;
            metroTabControl6.Enabled = false;
            metroTabControl7.Enabled = false;
            metroTabControl8.Enabled = false;
            metroTile3.Visible = false;
        }
        
        public static class WordExporter
        {
            /// <summary>
            /// Завершение экспорта
            /// </summary>
            internal static Action ExportCompleted;
            private static DataGridView _dataGridView;
            private static dynamic _wdApp;//Приложение Word
            private static string _fileName;//путь к файлу
            private static dynamic _wdTable;//Таблица в документе
            private static dynamic _wdDoc;//Документ
                                          /// <summary>
                                          /// Экспорт содержимого DataGridView в Word
                                          /// </summary>
                                          /// <param name="dgv">Экспортируемый DataGridView</param>
                                          /// <param name="filename">Имя файла</param>
            public static void Export(DataGridView dgv, string filename)
            {
                //Сохраняем ссылки на объекты
                _dataGridView = dgv;
                _fileName = filename;
                //Создаём новое приложение Word
                _wdApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                _wdApp.Visible = true;//Показываем
                                      //Поток, в котором будет проводиться экспорт
                var exportthread = new Thread(ExporterDelegate);
                //Начинаем экспорт
                exportthread.Start();
            }
            /// <summary>
            /// Делегат, выполняющий экспорт.
            /// </summary>
            private static void ExporterDelegate()
            {
                //Пытаемся создать новый документ
                try
                {
                    _wdDoc = _wdApp.Documents.Add();
                }
                catch (Exception ex)
                {
                    ExportCompleted?.Invoke();
                    throw new Exception($"Ошибка при создании документа{_fileName}.", ex);
                }
                //Создаём в документе таблицу и строку заголовка
                CreateWordTable();
                //Перебор строк
                foreach (DataGridViewRow row in _dataGridView.Rows)
                {
                    //На случай, если пользователю разрешено добавлять строки
                    //Новую строку пропускаем
                    if (row.IsNewRow) continue;
                    //Добавляем из строки dgv данные в строку таблицы документа
                    AddRow(row);
                }
                //Сохранение документа
                SaveDoc();
                //Закрываем документ
                _wdDoc.Close(false);
                //убираем ссылку на объект
                _wdDoc = null;
                //Выходим из приложения
                _wdApp.Quit();
                //убираем ссылку на объект
                _wdApp = null;
                //Если есть — вызываем метод окончания экспорта
                ExportCompleted?.Invoke();
            }
            /// <summary>
            /// Сохранение документа
            /// </summary>
            private static void SaveDoc()
            {
                //Формат сохранения файла в зависимости от расширения
                var fileFormat = _fileName.EndsWith("doc") ? 0 : 16;
                //Версия приложения, чтобы определить какой метод сохранения вызывать.
                var appVersion = float.Parse(_wdApp.Version.ToString(), CultureInfo.InvariantCulture);
                if (appVersion < 14)
                {
                    //Если приложение старше Word 2010
                    _wdDoc.SaveAs(_fileName, fileFormat, false, string.Empty, false);
                }
                else
                {
                    _wdDoc.SaveAs2(_fileName, fileFormat, false, string.Empty, false);
                }
            }
            /// <summary>
            /// Добавление строки в таблицу документа
            /// </summary>
            /// <param name="row">Строка из DataGridView, данные из которой нужно добавить в документ</param>
            private static void AddRow(DataGridViewRow row)
            {
                dynamic wdRow = _wdDoc.Tables[1].Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    wdRow.Cells[cell.OwningColumn.DisplayIndex + 1].Range.Text = cell.Value.ToString();
                }
            }
            /// <summary>
            /// Создание таблицы в документе
            /// </summary>
            /// <param name="dgv">DataGridView из которого создаётся таблица</param>
            private static void CreateWordTable()
            {
                _wdTable = _wdDoc.Tables.Add(_wdDoc.Range, 1, _dataGridView.ColumnCount);
                foreach (DataGridViewColumn column in _dataGridView.Columns)
                {
                    if (column.Index == -1) continue;
                    _wdTable.Cell(1, column.DisplayIndex + 1).Range.Text = column.HeaderText;
                }
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            metroTile3.Visible = true;
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = metroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = metroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(metroDateTime1.Text) < Convert.ToDateTime(metroDateTime2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
WHERE сотрудник.Дата Between #{0}# and #{1}#", date1, date2), conn);
                adapter.Fill(ds);
                metroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTabControl5.Enabled = false;
                metroTabControl7.Enabled = false;
                metroTabControl8.Enabled = false;
                metroTile10.Enabled = false;
                metroTile16.Enabled = false;
                metroTabControl8.Enabled = false;
                //EnableOFF();
                if (metroGrid1.RowCount == 0)
                {
                    MessageBox.Show("Запись не найдена", "Сотрудника не найдено");
                    upd();

                }
                //metroGrid1.Text = "";
            
        }
            else
            {
                MessageBox.Show("Начальный период не может быть больше конечного", "Ошибка диапазона");
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (textBox3.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                //panel3.Visible = false;
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника 
                           INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                           WHERE Фамилия='" + textBox3.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                metroGrid1.DataSource = dt;
                metroGrid1.Columns[8].Visible = false;
                metroGrid1.Columns[0].Visible = false;
                metroTabControl5.Enabled = false;
                metroTabControl6.Enabled = false;
                metroTile14.Enabled = false;
                metroTile15.Enabled = false;
                metroTabControl8.Enabled = false;
                metroTile3.Visible = true;

                if (metroGrid1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено", "Сотрудника не найдено");
                    upd();
                    
                }
                metroGrid1.Text = "";
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                // panel3.Visible = false;
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Имя='" + textBox4.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                metroGrid1.DataSource = dt;
               // EnableOFF();
                metroGrid1.Columns[8].Visible = false;
                metroGrid1.Columns[0].Visible = false;

                metroTabControl5.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTile14.Enabled = false;
                metroTile15.Enabled = false;
                metroTabControl8.Enabled = false;
                metroTile3.Visible = true;


                if (metroGrid1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено", "Сотрудника не найдено");
                    upd();
                    metroTile3.Visible = true;
                }
                textBox4.Text = "";
            }
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            metroTile3.Visible = true;
            if ((textBox5.Text == "") || (textBox8.Text == ""))
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                    string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Фамилия='" + textBox5.Text + "' and Имя='" + textBox8.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                metroGrid1.DataSource = dt;
                metroGrid1.Columns[8].Visible = false;
               // metroGrid1.Columns[0].Visible = false;

                metroTabControl4.Enabled = false;
                metroTabControl3.Enabled = false;
                metroTile14.Enabled = false;
                metroTile15.Enabled = false;
                metroTabControl8.Enabled = false;
                metroTile3.Visible = true;

                // EnableOFF();
                if (metroGrid1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено", "Сотрудника не найдено");
                    upd();

                }
                textBox5.Text = "";
                textBox8.Text = "";
            }
        }

        private void сброситьФильтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upd();
        }

        private void metroGrid1_SelectionChanged(object sender, EventArgs e)
        {
            if (metroGrid1.CurrentRow != null)
            {
               pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                metroTextBox1.Text = Convert.ToString(metroGrid1.CurrentRow.Cells[8].Value);
                pictureBox1.Load(Application.StartupPath + @"\model\" + metroTextBox1.Text);
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
            metroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            if (metroGrid1.RowCount == 0)
            {
                MessageBox.Show("Запрос не дал результатов", "Запрос пуст");
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
           metroGrid1.DataSource = ds.Tables[0];
            conn.Close();
            EnableOFF();
            metroTabControl3.Enabled = false;
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
            metroGrid1.DataSource = ds.Tables[0];
            metroGrid1.Columns[8].Visible = false;
            metroGrid1.Columns[7].Visible = false;
            metroGrid1.Columns[6].Visible = false;
            metroGrid1.Columns[5].Visible = false;

            conn.Close();
            EnableOFF();
            metroTabControl3.Enabled = false;
            if (metroGrid1.RowCount == 0)
            {
                MessageBox.Show("Запрос не дал результатов", "Запрос пуст");
                upd();

            }
        }

        private void metroTile7_Click(object sender, EventArgs e)
        {
            Money main2 = new Money();
            //   this.Opacity = 0.6; // прозрачность формы
            //  main2.Owner = this; // для доступа к главной форме из других
            main2.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            metroTile3.Visible = true;
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
                if (Convert.ToInt32(metroTextBox6.Text) < Convert.ToInt32(metroTextBox7.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата where зарплата_сотрудника.зарплата between {0} and {1}", metroTextBox6.Text, metroTextBox7.Text), conn);
                adapter.Fill(ds);
                metroGrid1.DataSource = ds.Tables[0];
                conn.Close();
                metroTile10.Enabled = false;
                metroTile11.Enabled = false;
                metroTabControl5.Enabled = false;
                metroTabControl7.Enabled = false;
                metroTabControl8.Enabled = false;

                if (metroGrid1.RowCount == 0)
                {
                    MessageBox.Show("Запись не найдена", "Сотрудника не найдено");
                    upd();

                }
            }
            else
            {
                MessageBox.Show("Начальная зарплата не может быть больше конечной", "Ошибка диапазона");
                upd();
            }


            
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void metroTile14_Click(object sender, EventArgs e)
        {

        }

        private void metroGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void metroTile21_Click(object sender, EventArgs e)
        {
            Send_Mail send = new Send_Mail();
            send.ShowDialog();
        }
    }
}