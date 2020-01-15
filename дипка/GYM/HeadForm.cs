using System;
using System.Collections.Generic;
using System.ComponentModel;
using Word = Microsoft.Office.Interop.Word;
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
    public partial class HeadForm : Form
    {
        public HeadForm()
        {
            InitializeComponent();
        }
        private DS_Money.Зарплата_сотрудникаDataTable dta1 { get; set; }
        public int rez = 0;
        public string id;
        public DataTable dt;
        string filename;
        public OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
        public OleDbDataAdapter sda;

        public void upd()
          {
            sda = new OleDbDataAdapter(@"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата  
                                        FROM Зарплата_сотрудника 
                                        INNER JOIN Сотрудник 
                                        ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", con);
            
            dt = new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;
            panel3.Visible = true;
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;
            groupBox3.Enabled = true;
            groupBox4.Enabled = true;
            groupBox5.Enabled = true;
            groupBox6.Enabled = true;
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Select();
            dataGridView1.AllowUserToAddRows = false;
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
            dataGridView1.DataSource = dt;
        }

        private void HeadForm_Load(object sender, EventArgs e)
        {
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

        private void button1_Click(object sender, EventArgs e)
        {
            AddEmployee ObjEmployeeAdd = new AddEmployee();
            ObjEmployeeAdd.textBox1.Text = "";
            ObjEmployeeAdd.textBox2.Text = "";
            ObjEmployeeAdd.textBox3.Text = "";
            ObjEmployeeAdd.maskedTextBox1.Text = "";
            ObjEmployeeAdd.dateTimePicker1.Text = null;
            ObjEmployeeAdd.textBox6.Text = "";
            ObjEmployeeAdd.textBox5.Text = "";
            ObjEmployeeAdd.Text = "Добавить сотрудника";
            ObjEmployeeAdd.button1.Text = "Добавить";
            if (ObjEmployeeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                   

                        dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                        id = Convert.ToString(Convert.ToInt32(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value) + 1);
                        con.Open();
                        OleDbCommand sss = new OleDbCommand(@"INSERT INTO [Сотрудник]
                                                        ( Фамилия, Имя, Отчество, Должность,Телефон, Дата,Пароль,Фото,Идзарплата)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5,@st6,@st7,@st8,@st9)", con);
                        sss.Parameters.AddWithValue("st1", ObjEmployeeAdd.textBox1.Text);
                        sss.Parameters.AddWithValue("st2", ObjEmployeeAdd.textBox2.Text);
                        sss.Parameters.AddWithValue("st3", ObjEmployeeAdd.textBox3.Text);
                        sss.Parameters.AddWithValue("st4", ObjEmployeeAdd.textBox4.Text);
                        sss.Parameters.AddWithValue("st5", ObjEmployeeAdd.maskedTextBox1.Text);
                        sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjEmployeeAdd.dateTimePicker1.Text));
                        sss.Parameters.AddWithValue("st7", Convert.ToInt16(ObjEmployeeAdd.textBox5.Text));
                        sss.Parameters.AddWithValue("st8", ObjEmployeeAdd.textBox6.Text);
                        sss.Parameters.AddWithValue("st9", Convert.ToInt16(ObjEmployeeAdd.comboBox1.SelectedValue.ToString()));
                        sss.ExecuteNonQuery();
                        con.Close();
                        upd();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }

        private void button2_Click(object sender, EventArgs e)
        {
              AddEmployee ObjEmployeeUpdate = new AddEmployee();
              //  int t = 0;
              //t = Convert.ToInt32(dataGridView1.CurrentRow.Cells[10].Value);
              //  //   DELETE FROM сотрудник WHERE идсотрудник = " + rez + "", con
              //  string s = "SELECT  идЗарплата  FROM зарплата_сотрудника WHERE идЗарплата =" + t + "";
              //  sda = new OleDbDataAdapter(s, con);
              //  dt = new DataTable();
              //  sda.Fill(dt);
                ObjEmployeeUpdate.Text = "Редактировать сотрудника";
                ObjEmployeeUpdate.button1.Text = "Редактировать";
                rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                ObjEmployeeUpdate.textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
                ObjEmployeeUpdate.textBox2.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[2].Value);
                ObjEmployeeUpdate.textBox3.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value);
                ObjEmployeeUpdate.textBox4.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value);
                ObjEmployeeUpdate.maskedTextBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[5].Value);
                ObjEmployeeUpdate.dateTimePicker1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[6].Value);
                ObjEmployeeUpdate.textBox5.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[7].Value);
                ObjEmployeeUpdate.textBox6.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[8].Value);
                 //  ObjEmployeeUpdate.comboBox1.ValueMember = Convert.ToString( dt);
                ObjEmployeeUpdate.comboBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[9].Value);
                if (ObjEmployeeUpdate.ShowDialog() == DialogResult.OK)
                    try
                    {
                        dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                        con.Open();
                        OleDbCommand sss = new OleDbCommand("update Сотрудник set Фамилия=@st1, Имя=@st2, Отчество=@st3, Должность=@st4,Телефон=@st5, Дата=@st6,Пароль=@st7,Фото=@st8,Идзарплата=@st9 where идсотрудник=" + rez + "", con);
                        sss.Parameters.AddWithValue("st1", ObjEmployeeUpdate.textBox1.Text);
                        sss.Parameters.AddWithValue("st2", ObjEmployeeUpdate.textBox2.Text);
                        sss.Parameters.AddWithValue("st3", ObjEmployeeUpdate.textBox3.Text);
                        sss.Parameters.AddWithValue("st4", ObjEmployeeUpdate.textBox4.Text);
                        sss.Parameters.AddWithValue("st5", ObjEmployeeUpdate.maskedTextBox1.Text);
                        sss.Parameters.AddWithValue("st6", Convert.ToDateTime(ObjEmployeeUpdate.dateTimePicker1.Text));
                        sss.Parameters.AddWithValue("st7", Convert.ToInt32(ObjEmployeeUpdate.textBox5.Text));
                        sss.Parameters.AddWithValue("st8", ObjEmployeeUpdate.textBox6.Text);
                        sss.Parameters.AddWithValue("st9", Convert.ToInt32(ObjEmployeeUpdate.comboBox1.SelectedValue.ToString()));
                        sss.ExecuteNonQuery();
                        con.Close();
                        upd();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Вы уверены, что хотите Удалить?", "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                con.Open();
                rez = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                OleDbCommand sss = new OleDbCommand(@"DELETE FROM сотрудник 
                                                    WHERE идсотрудник=" + rez + "", con);
                sss.ExecuteNonQuery();
                upd();
                con.Close();
            }
        }

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[8].Value);
            pictureBox1.Load(Application.StartupPath + @"\model\" + textBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
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
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }
                ExcelApp.Visible = true;
                ExcelApp.UserControl = true;
            }
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
        private void button5_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(delegate (object obj)
            {
                MessageBox.Show("Выполняется формирование отчета, это займёт около 10 секунд");
            }).Start();
            //Переключение активности кнопки
            var setButtonEnabled = new Action<bool>(enabled => button5.Enabled = enabled);
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
            WordExporter.Export(dataGridView1, Path.GetFullPath("exported.doc"));
        }

        private void textBox2_KeyUp(object sender, KeyEventArgs e)
        {
            panel3.Visible = true;
            string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                       FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                       WHERE[Фамилия] LIKE '%" + textBox2.Text + "%'";
            sda = new OleDbDataAdapter(s, con);
            dt = new DataTable();
            sda.Fill(dt);
            dataGridView1.DataSource = dt;
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("Таких нет");
                textBox2.Text = "";
                upd();
            }
        }
        public void EnabledOFF()
        {
           // panel3.Visible = false;
            groupBox2.Enabled = false;
            groupBox3.Enabled = false;
            groupBox4.Enabled = false;
            groupBox5.Enabled = false;
            groupBox6.Enabled = false;
        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            panel3.Visible = true;
            string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ISgym.mdb");
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();
            DataSet ds = new DataSet();
            string date1 = dateTimePicker1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            string date2 = dateTimePicker2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
            if (Convert.ToDateTime(dateTimePicker1.Text) < Convert.ToDateTime(dateTimePicker2.Text))
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format(@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
WHERE сотрудник.Дата Between #{0}# and #{1}#", date1, date2), conn);


                                                                              
                adapter.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                conn.Close();
                EnabledOFF();
             }
               else
              {
                MessageBox.Show("Начальный период не может быть больше конечного", "Ошибка диапазона");
                upd();
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
                    DialogResult result = MessageBox.Show("Неверный тип данных", "Корректность ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Length == 1)
                ((TextBox)sender).Text = ((TextBox)sender).Text.ToUpper();
                ((TextBox)sender).Select(((TextBox)sender).Text.Length, 0);
        }

        private void button7_Click(object sender, EventArgs e)
        {
             if (textBox4.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                //panel3.Visible = false;
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника 
                           INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                           WHERE Фамилия='" + textBox4.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[0].Visible = false;
                EnabledOFF();
                if (dataGridView1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено, Сотрудника не найдено");
                    upd();
                    panel3.Visible = true;
                }
                    textBox4.Text = "";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox5.Text == "")
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
               // panel3.Visible = false;
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Имя='" + textBox5.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
                EnabledOFF();
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[0].Visible = false;
                if (dataGridView1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено, Сотрудника не найдено");
                    upd(); panel3.Visible = true;
                }
                   textBox5.Text = "";
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel3.Visible = true;
            if ((textBox3.Text == "") || (textBox6.Text == ""))
            {
                MessageBox.Show("Не введены данные");
            }
            else
            {
                string s = @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Фамилия='" + textBox6.Text + "' and Имя='" + textBox3.Text + "'";
                sda = new OleDbDataAdapter(s, con);
                dt = new DataTable();
                sda.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[8].Visible = false;
                dataGridView1.Columns[0].Visible = false;
                EnabledOFF();
                if (dataGridView1.RowCount == 0)
                {
                    MessageBox.Show("Таких сотрудников не найдено, Сотрудника не найдено");
                    upd();
                }
                textBox6.Text = "";
                textBox3.Text = "";
            }
        }

        private void contextMenuStrip1_Opening_1(object sender, CancelEventArgs e)
        {
            foreach (ToolStripMenuItem item in contextMenuStrip1.Items)
                Sub(item, cms_click);
        }

            private void Sub(ToolStripMenuItem item, EventHandler eventHandler)
        {
            item.Click += eventHandler;
            //Подменю
            foreach (ToolStripMenuItem subItem in item.DropDownItems)
                Sub(subItem, eventHandler);
        }

        void cms_click(object sender, EventArgs e)
        {
            upd(); 
        }

       private void dataGridView1_SelectionChanged_1(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[8].Value);
                pictureBox1.Load(Application.StartupPath + @"\model\" + textBox1.Text);
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
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
                dataGridView1.DataSource = ds.Tables[0];
                conn.Close();
                EnabledOFF();
              groupBox1.Enabled = false;
            }

        private void button12_Click(object sender, EventArgs e)
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
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            EnabledOFF();
            groupBox1.Enabled = false;
        }

        private void button11_Click(object sender, EventArgs e)
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
            dataGridView1.DataSource = ds.Tables[0];
            conn.Close();
            EnabledOFF();
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[8].Visible = false;
            groupBox1.Enabled = false;
        }

        private void button13_Click(object sender, EventArgs e)
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
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[8].Visible = false;
            dataGridView1.Columns[7].Visible = false;
            dataGridView1.Columns[6].Visible = false;
            dataGridView1.Columns[5].Visible = false;

            conn.Close();
            EnabledOFF();
            groupBox1.Enabled = false;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Report rep = new Report();
            rep.Show();
            rep.reportViewer1.Visible = true;
        }

        private void button16_Click(object sender, EventArgs e)
        {

            Money main2 = new Money();
         //   this.Opacity = 0.6; // прозрачность формы
          //  main2.Owner = this; // для доступа к главной форме из других
            main2.ShowDialog();


        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox4_TextChanged(object sender, EventArgs e)
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

        private void textBox6_TextChanged(object sender, EventArgs e)
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
    }
    } 







