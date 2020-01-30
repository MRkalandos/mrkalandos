using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using MetroFramework;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;
using TextBox = System.Windows.Forms.TextBox;

namespace GYM
{
    public partial class HeadForm : MetroFramework.Forms.MetroForm
    {
        private const string Message = @"Неверный тип данных";
        private const string Title = @"Корректность ввода";
        public int idSportsmen;
        public int idTrenerovka;
        public int idEmployee;
        public int idTrener;
        public int idSale;
        public int idAbonement;
        public DataTable dataTableEmployee;
        public DataTable dataTableSportsmen;
        public DataTable dataTableSale;
        public DataTable dataTableAbonement;
        public DataTable dataTableTrener;
        public DataTable dataTableTrenerovka;
        string _filename;

        public OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                                Directory.GetParent(Directory.GetCurrentDirectory())
                                                                    .Parent?.FullName +
                                                                "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                   Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
                                   "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public OleDbDataAdapter dataAdapterEmployee;
        public OleDbDataAdapter dataAdapterSale;
        public OleDbDataAdapter dataAdapterSportsmen;
        public OleDbDataAdapter dataAdapterAbonement;
        public OleDbDataAdapter dataAdapterTrener;
        public OleDbDataAdapter dataAdapterTrening;
        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/HeadForm.txt";

        public HeadForm()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        public void UpdSportsmen()
        {
            try
            {
                dataAdapterSportsmen = new OleDbDataAdapter(@"SELECT Спортсмен.*
                                                                          FROM Спортсмен;", connection);
                dataTableSportsmen = new DataTable();
                dataAdapterSportsmen.Fill(dataTableSportsmen);
                SPORTMmetroGrid2.DataSource = dataTableSportsmen;
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void UpdTrening()
        {
            try
            {
                dataAdapterTrening = new OleDbDataAdapter(
                    @"SELECT Тренировка.Идтренировка, Тренировка.Название as [Название тренировки], Вид_тренировки.Название as [Название вида], Тренер.Фамилия as [Тренер],тренер.идтренер, вид_тренировки.идвидтренировка
                                 FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка 
                                 ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) 
                                 ON Тренер.Идтренер = Тренировка.Идтренер;", connection);
                dataTableTrenerovka = new DataTable();
                dataAdapterTrening.Fill(dataTableTrenerovka);
                TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
                TRENINGmetroGrid1.Columns[0].Visible = false;
                TRENINGmetroGrid1.Columns[4].Visible = false;
                TRENINGmetroGrid1.Columns[5].Visible = false;
                TRENINGmetroGrid1.Select();
                TRENINGmetroGrid1.AllowUserToAddRows = false;
                TRENINGmetroTabControl6.Enabled = true;
                //metroTile18.Enabled = true;
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void UpdSale()
        {
            try
            {
                dataAdapterSale = new OleDbDataAdapter(
                    @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия as [Спортсмен], Абонемент.Название, Абонемент.Количество_посещений as [Количество посещений], Продажа_абонемента.Дата_начала as [Дата начала], Продажа_абонемента.Дата_окончания as [Дата окончания], Сотрудник.Фамилия as [Сотрудник], Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
                                 FROM Сотрудник 
                                 INNER JOIN (Абонемент INNER JOIN (Спортсмен INNER JOIN Продажа_абонемента 
                                 ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
                                 ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент)
                                 ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник;", connection);
                dataTableSale = new DataTable();
                dataAdapterSale.Fill(dataTableSale);
                SALEmetroGrid1.DataSource = dataTableSale;
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void UpdAbonement()
        {
            try
            {
                dataAdapterAbonement = new OleDbDataAdapter(
                    @"SELECT Абонемент.Идабонемент, Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений as[Количество посещений], Тренировка.Название, Абонемент.Идтренеровка
                                 FROM Тренировка INNER JOIN Абонемент
                                 ON Тренировка.Идтренировка = Абонемент.Идтренеровка;", connection);
                dataTableAbonement = new DataTable();
                dataAdapterAbonement.Fill(dataTableAbonement);
                ABONmetroGrid1.DataSource = dataTableAbonement;
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void UpdTrener()
        {
            try
            {
                dataAdapterTrener = new OleDbDataAdapter(@"SELECT Тренер.*
                                                                      FROM Тренер;", connection);
                dataTableTrener = new DataTable();
                dataAdapterTrener.Fill(dataTableTrener);
                TRENmetroGrid1.DataSource = dataTableTrener;
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void UpdEmployee()
        {
            try
            {
                dataAdapterEmployee = new OleDbDataAdapter(
                    @"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата, сотрудник.идзарплата  
                                        FROM Зарплата_сотрудника 
                                        INNER JOIN Сотрудник 
                                        ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", connection);
                dataTableEmployee = new DataTable();
                dataAdapterEmployee.Fill(dataTableEmployee);
                EMPLmetroGrid1.DataSource = dataTableEmployee;
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
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void Updfilename()
        {
            var connectionEmplyee =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterEmployee = new OleDbDataAdapter(
                @"SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Сотрудник.Фото, Зарплата_сотрудника.Зарплата
                                       FROM Зарплата_сотрудника 
                                       INNER JOIN Сотрудник   
                                       ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата;", connectionEmplyee);
            dataTableEmployee = new DataTable();
            dataAdapterEmployee.Fill(dataTableEmployee);
            EMPLmetroGrid1.DataSource = dataTableEmployee;
            var connectionSportsmen =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterSportsmen = new OleDbDataAdapter(@"SELECT Спортсмен.*
                                                                      FROM Спортсмен;", connectionSportsmen);
            dataTableSportsmen = new DataTable();
            dataAdapterSportsmen.Fill(dataTableSportsmen);
            SPORTMmetroGrid2.DataSource = dataTableSportsmen;
            var conAbonement =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterAbonement = new OleDbDataAdapter(
                @"SELECT Абонемент.Идабонемент, Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                             FROM Тренировка INNER JOIN Абонемент
                             ON Тренировка.Идтренировка = Абонемент.Идтренеровка;", conAbonement);
            dataTableAbonement = new DataTable();
            dataAdapterAbonement.Fill(dataTableAbonement);
            ABONmetroGrid1.DataSource = dataTableAbonement;
            var conSale =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterAbonement = new OleDbDataAdapter(
                @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
                             FROM Сотрудник INNER JOIN (Абонемент 
                             INNER JOIN (Спортсмен 
                             INNER JOIN Продажа_абонемента 
                             ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) 
                             ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                             ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник;", conSale);
            dataTableSale = new DataTable();
            dataAdapterSale.Fill(dataTableSale);
            SALEmetroGrid1.DataSource = dataTableSale;
            var contrenerovka =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterTrening = new OleDbDataAdapter(
                @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия
                             FROM Тренер 
                             INNER JOIN (Вид_тренировки 
                             INNER JOIN Тренировка 
                             ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) 
                             ON Тренер.Идтренер = Тренировка.Идтренер;", contrenerovka);
            dataTableTrenerovka = new DataTable();
            dataAdapterTrening.Fill(dataTableTrenerovka);
            TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
            var conTrener =
                new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _filename + "");
            dataAdapterTrener = new OleDbDataAdapter(@"SELECT тренер.*
                                                 FROM тренер;", conTrener);
            dataTableTrener = new DataTable();
            dataAdapterTrener.Fill(dataTableTrener);
            TRENmetroGrid1.DataSource = dataTableTrener;
            TRENmetroGrid1.Columns[0].Visible = false;
            SPORTMmetroGrid2.Columns[0].Visible = false;
            EMPLmetroGrid1.Columns[0].Visible = false;
            TRENINGmetroGrid1.Columns[0].Visible = false;
        }

        private void HeadeForm_Load(object sender, EventArgs e)
        {
            SPORTMmetroComboBox1.SelectedIndex = 0;
            FocusMe();
            try
            {
                UpdTrener();
                UpdEmployee();
                UpdSportsmen();
                UpdTrening();
                UpdAbonement();
                UpdSale();
            }
            catch
            {
                MetroMessageBox.Show(this, "\nФайл базы данных не найден укажите новый путь",
                    "Файл БД не найден", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var openFileDialog1 = new OpenFileDialog() {Filter = @"Файл БД (*.mdb)|*.mdb"};
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (var sw =
                        File.CreateText(Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName + @"\" +
                                        "text.txt"))
                    {
                        _filename = openFileDialog1.FileName;
                        sw.WriteLine(_filename);
                        File.Copy(_filename,
                            Directory.GetParent(Directory.GetCurrentDirectory()).Parent
                                ?.FullName /*AppDomain.CurrentDomain.BaseDirectory*/ + Path.GetFileName(_filename));
                        Updfilename();
                    }
                }
                else
                {
                    string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName + @"\" +
                                  "text.txt";
                    using (var sr = new StreamReader(path))
                    {
                        _filename = sr.ReadLine();
                    }

                    Updfilename();
                    MetroMessageBox.Show(this,
                        "\nВы не выбрали файл БД, поэтому выбранно последнее выбранное место БД", "Не указана БД",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void metroTile18_Click(object sender, EventArgs e)
        {
            try
            {
                var dbConnection = new OleDbConnection(conString);
                dbConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECt сотрудник.идсотрудник, Сотрудник.Фамилия,Сотрудник.Имя,Сотрудник.отчество,Сотрудник.Дата,Сотрудник.Должность,Сотрудник.Телефон, '(' & Зарплата_сотрудника.Зарплата & ')' AS Зарплата ,сотрудник.фото
                                                             FROM Зарплата_сотрудника 
                                                             INNER JOIN Сотрудник 
                                                             ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                                                             UNION 
                                                             SELECT 'Средняя:','', '','','','','', Avg(Зарплата_сотрудника.Зарплата) AS Средняя,'' 
                                                             FROM Зарплата_сотрудника INNER JOIN Сотрудник 
                                                             ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата ",
                    dbConnection);
                adapter.Fill(dataSet);
                EMPLmetroGrid1.DataSource = dataSet.Tables[0];
                dbConnection.Close();
                EnableOff();
                EMPLmetroGrid1.Columns[0].Visible = false;
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroTabControl3.Enabled = false;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile4_Click(object sender, EventArgs e)
        {
            var objEmployeeAdd = new ModEmployee
            {
                textBox1 = {Text = ""},
                textBox2 = {Text = ""},
                textBox3 = {Text = ""},
                maskedTextBox1 = {Text = ""},
                metroDateTime1 = {Text = null},
                metroTextBox5 = {Text = ""},
                label1= { Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[0].Value)},
                Text = @"Добавить сотрудника",
                metroTile1 = {Text = @"Добавить"}
            };
            connection.Close();
            connection.Open();
            var cmd = new OleDbCommand("select идзарплата,зарплата from зарплата_сотрудника", connection);
            objEmployeeAdd.metroComboBox1.DisplayMember = "Зарплата";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, int>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (int) reader[1]);
            }
            reader.Close();
            cmd.ExecuteNonQuery();
            objEmployeeAdd.metroComboBox1.DataSource = list.ToList();
            objEmployeeAdd.metroComboBox1.DisplayMember = "Value";
            objEmployeeAdd.metroComboBox1.ValueMember = "Key";
            connection.Close();
            if (objEmployeeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    EMPLmetroGrid1.Sort(EMPLmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddEmployee = new OleDbCommand(@"INSERT INTO [Сотрудник]
                                                        ( Фамилия, Имя,  Отчество, Должность, Телефон, Дата, Пароль,Фото,  Идзарплата)
                                                  VALUES(@surname, @name,@patryon, @position, @tel,    @date,@pas,  @photo,@idmoney)",
                        connection);
                    queryAddEmployee.Parameters.AddWithValue("surname", objEmployeeAdd.textBox1.Text);
                    queryAddEmployee.Parameters.AddWithValue("name", objEmployeeAdd.textBox2.Text);
                    queryAddEmployee.Parameters.AddWithValue("patryon", objEmployeeAdd.textBox3.Text);
                    queryAddEmployee.Parameters.AddWithValue("position", objEmployeeAdd.metroTextBox4.Text);
                    queryAddEmployee.Parameters.AddWithValue("tel", objEmployeeAdd.maskedTextBox1.Text);
                    queryAddEmployee.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objEmployeeAdd.metroDateTime1.Text));
                    queryAddEmployee.Parameters.AddWithValue("pas", Convert.ToInt32(objEmployeeAdd.metroTextBox5.Text));
                    queryAddEmployee.Parameters.AddWithValue("photo", objEmployeeAdd.metroTextBox6.Text);
                    queryAddEmployee.Parameters.AddWithValue("idmoney",
                        Convert.ToInt32(objEmployeeAdd.metroComboBox1.SelectedValue.ToString()));
                    queryAddEmployee.ExecuteNonQuery();
                    connection.Close();
                    UpdEmployee();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            var objEmployeeUpdate = new ModEmployee();
            connection.Close();
            objEmployeeUpdate.Text = @"Редактировать сотрудника";
            objEmployeeUpdate.metroTile1.Text = @"Редактировать";
            Debug.Assert(EMPLmetroGrid1.CurrentRow != null, "Таблица пуста");
            idEmployee = Convert.ToInt32(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
            objEmployeeUpdate.label1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
            objEmployeeUpdate.textBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[1].Value);
            objEmployeeUpdate.textBox2.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[2].Value);
            objEmployeeUpdate.textBox3.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[3].Value);
            objEmployeeUpdate.metroTextBox4.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[4].Value);
            objEmployeeUpdate.maskedTextBox1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[5].Value);
            objEmployeeUpdate.metroDateTime1.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[6].Value);
            objEmployeeUpdate.metroTextBox5.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[7].Value);
            objEmployeeUpdate.metroTextBox6.Text = Convert.ToString(EMPLmetroGrid1.CurrentRow.Cells[8].Value);
            connection.Open();
            var cmd = new OleDbCommand("select идзарплата,зарплата from зарплата_сотрудника", connection);
            objEmployeeUpdate.metroComboBox1.DisplayMember = "Зарплата";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, int>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (int) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            connection.Close();
            objEmployeeUpdate.metroComboBox1.DataSource = list.ToList();
            objEmployeeUpdate.metroComboBox1.DisplayMember = "Value";
            objEmployeeUpdate.metroComboBox1.ValueMember = "Key";
            objEmployeeUpdate.metroComboBox1.SelectedValue = EMPLmetroGrid1.CurrentRow.Cells[10].Value;
            if (objEmployeeUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    EMPLmetroGrid1.Sort(EMPLmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateEmployee =
                        new OleDbCommand(
                            "update Сотрудник set Фамилия=@name, Имя=@surname, Отчество=@patryon, Должность=@position,Телефон=@telephon, Дата=@date,Пароль=@pass,Фото=@photo,Идзарплата=@idmoney where идсотрудник=" +
                            idEmployee + "", connection);
                    queryUpdateEmployee.Parameters.AddWithValue("name", objEmployeeUpdate.textBox1.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("surname", objEmployeeUpdate.textBox2.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("patryon", objEmployeeUpdate.textBox3.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("position", objEmployeeUpdate.metroTextBox4.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("telephon", objEmployeeUpdate.maskedTextBox1.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objEmployeeUpdate.metroDateTime1.Text));
                    queryUpdateEmployee.Parameters.AddWithValue("pass",
                        Convert.ToInt32(objEmployeeUpdate.metroTextBox5.Text));
                    queryUpdateEmployee.Parameters.AddWithValue("photo", objEmployeeUpdate.metroTextBox6.Text);
                    queryUpdateEmployee.Parameters.AddWithValue("idmoney",
                        Convert.ToInt32(objEmployeeUpdate.metroComboBox1.SelectedValue));
                    queryUpdateEmployee.ExecuteNonQuery();
                    connection.Close();
                    UpdEmployee();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    connection.Close();
                    connection.Open();
                   // Debug.Assert(EMPLmetroGrid1.CurrentRow != null, "Таблица пуста");
                   idEmployee = Convert.ToInt32(EMPLmetroGrid1.CurrentRow.Cells[0].Value);
                    var queryDeleteEmployee = new OleDbCommand(@"DELETE FROM сотрудник 
                                                    WHERE идсотрудник=" + idEmployee + "", connection);
                    queryDeleteEmployee.ExecuteNonQuery();
                    connection.Close();
                    UpdEmployee();
                    
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile8_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Employee.xlsx";
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.UpdEmployee();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var excelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            var excelWorkSheet = (Worksheet) excelWorkBook.Worksheets.get_Item(1);
            excelWorkSheet.StandardWidth = 17;
            excelWorkSheet.Name = "Сотрудники";
            excelApp.Cells[1, 1] = "Фамилия";
            excelApp.Cells[1, 2] = "Имя";
            excelApp.Cells[1, 3] = "Отчество";
            excelApp.Cells[1, 4] = "Должность";
            excelApp.Cells[1, 5] = "Телефон";
            excelApp.Cells[1, 6] = "Дата";
            excelApp.Cells[1, 7] = "Пароль";
            excelApp.Cells[1, 8] = "Зарплата";
            {
                for (int i = 1; i < EMPLmetroGrid1.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = EMPLmetroGrid1.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = EMPLmetroGrid1.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = EMPLmetroGrid1.Rows[i - 1].Cells[3].Value;
                    excelApp.Cells[i + 1, 4] = EMPLmetroGrid1.Rows[i - 1].Cells[4].Value;
                    excelApp.Cells[i + 1, 5] = EMPLmetroGrid1.Rows[i - 1].Cells[5].Value;
                    excelApp.Cells[i + 1, 6] = EMPLmetroGrid1.Rows[i - 1].Cells[6].Value;
                    excelApp.Cells[i + 1, 7] = EMPLmetroGrid1.Rows[i - 1].Cells[7].Value;
                    excelApp.Cells[i + 1, 8] = EMPLmetroGrid1.Rows[i - 1].Cells[9].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                    excelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    excelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEmailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по сотрудникам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEmailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по сотрудникам", Body = "Отчет по сотрудникам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587)
                            {
                                Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                EnableSsl = true
                            };
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                            objEmailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(
                                        new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                }
                            }
                        }
                }
            }
        }

        private void metroTile9_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Employee.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application {Visible = true};
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                var table = doc.Tables.Add(range, EMPLmetroGrid1.RowCount + 1, 8);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя";
                table.Cell(1, 3).Range.Text = "Отчество";
                table.Cell(1, 4).Range.Text = "Должность";
                table.Cell(1, 5).Range.Text = "Телефон";
                table.Cell(1, 6).Range.Text = "Дата";
                table.Cell(1, 7).Range.Text = "Пароль";
                table.Cell(1, 8).Range.Text = "Зарплата";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по сотрудникам", metroTile2 = {Text = @"Отправить"}
                        };
                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var from = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var to = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var messageObj = new MailMessage(from, to)
                                {
                                    Subject = "Отчет по сотрудникам",
                                    Body = "Отчет по сотрудникам",
                                    IsBodyHtml = true
                                };
                                messageObj.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587);
                                smtp.Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq");
                                smtp.EnableSsl = true;
                                smtp.Send(messageObj);
                                MessageBox.Show(@"Сообщение отправлено на почту" + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this,
                                    "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                    "Не отправлено",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                FocusMe();
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                            }
                    }
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                EMPLmetroTabControl5.Enabled = false;
                EMPLmetroTile11.Enabled = false;
                EMPLmetroTabControl7.Enabled = false;
                EMPLmetroTabControl8.Enabled = false;
                EMPLmetroTile16.Enabled = false;
                EMPLmetroTile3.Visible = true;
                string queryFindSurnameEmploye =
                    @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                       FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                       WHERE[Фамилия] LIKE '%" + EMPLtextBox2.Text + "%'";
                dataAdapterEmployee = new OleDbDataAdapter(queryFindSurnameEmploye, connection);
                dataTableEmployee = new DataTable();
                dataAdapterEmployee.Fill(dataTableEmployee);
                EMPLmetroGrid1.DataSource = dataTableEmployee;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    EMPLtextBox2.Text = "";
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        public void EnableOff()
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
            try
            {
                EMPLmetroTile3.Visible = true;
                var conn = new OleDbConnection(conString);
                conn.Open();
                var ds = new DataSet();
                string date1 = EMPLmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                string date2 = EMPLmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                if (Convert.ToDateTime(EMPLmetroDateTime1.Text) < Convert.ToDateTime(EMPLmetroDateTime2.Text))
                {
                    var adapter = new OleDbDataAdapter(
                        $@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                                  FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                                  WHERE сотрудник.Дата Between #{date1}# and #{date2}#", conn);
                    adapter.Fill(ds);
                    EMPLmetroGrid1.DataSource = ds.Tables[0];
                    conn.Close();
                    EMPLmetroTabControl5.Enabled = false;
                    EMPLmetroTabControl7.Enabled = false;
                    EMPLmetroTabControl8.Enabled = false;
                    EMPLmetroTile10.Enabled = false;
                    EMPLmetroTile16.Enabled = false;
                    EMPLmetroTabControl8.Enabled = false;
                    if (EMPLmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdEmployee();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного",
                        TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdEmployee();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                     MetroMessageBox.Show(this,Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void metroTextBox3_Click(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void metroTextBox4_Click(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            if (EMPLtextBox3.Text == "")
            {
                MetroMessageBox.Show(this,@"Заполните пустые(ое) поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string queryFindSurnameEmployee =
                        @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                      FROM Зарплата_сотрудника 
                      INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                      WHERE Фамилия='" + EMPLtextBox3.Text + "'";
                    dataAdapterEmployee = new OleDbDataAdapter(queryFindSurnameEmployee, connection);
                    dataTableEmployee = new DataTable();
                    dataAdapterEmployee.Fill(dataTableEmployee);
                    EMPLmetroGrid1.DataSource = dataTableEmployee;
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
                        MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdEmployee();
                    }

                    EMPLtextBox3.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if (EMPLtextBox4.Text == "")
            {
                MetroMessageBox.Show(this, "\nНе введены данные", "Корректность", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string queryFindNameEmployee =
                        @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Имя='" + EMPLtextBox4.Text + "'";
                    dataAdapterEmployee = new OleDbDataAdapter(queryFindNameEmployee, connection);
                    dataTableEmployee = new DataTable();
                    dataAdapterEmployee.Fill(dataTableEmployee);
                    EMPLmetroGrid1.DataSource = dataTableEmployee;
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
                        MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdEmployee();
                        EMPLmetroTile3.Visible = true;
                    }

                    EMPLtextBox4.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            EMPLmetroTile3.Visible = true;
            if ((EMPLtextBox5.Text == "") || (EMPLtextBox8.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе введены данные", TitleException,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string queryFindNameandSurname =
                        @"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                           FROM Зарплата_сотрудника INNER JOIN Сотрудник ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата
                           WHERE Фамилия='" + EMPLtextBox5.Text + "' and Имя='" + EMPLtextBox8.Text + "'";
                    dataAdapterEmployee = new OleDbDataAdapter(queryFindNameandSurname, connection);
                    dataTableEmployee = new DataTable();
                    dataAdapterEmployee.Fill(dataTableEmployee);
                    EMPLmetroGrid1.DataSource = dataTableEmployee;
                    EMPLmetroGrid1.Columns[8].Visible = false;
                    EMPLmetroTabControl4.Enabled = false;
                    EMPLmetroTabControl3.Enabled = false;
                    EMPLmetroTile14.Enabled = false;
                    EMPLmetroTile15.Enabled = false;
                    EMPLmetroTabControl8.Enabled = false;
                    EMPLmetroTile3.Visible = true;
                    if (EMPLmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Сотрудника не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdEmployee();
                    }

                    EMPLtextBox5.Text = "";
                    EMPLtextBox8.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        public void сброситьФильтрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                UpdEmployee();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
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
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile12_Click(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @" select Фамилия, Имя, Отчество, Должность, Телефон, Дата,Пароль,Идзарплата ,Фото
                                                                              FROM сотрудник 
                                                                              WHERE (сотрудник.[Дата] Is Not Null) 
                                                                              AND ((DateDiff('y',Date(),DateAdd('yyyy',DateDiff('yyyy',[Дата],Date()),[Дата])) 
                                                                              Between 0 And 31))", selectConnection);
                adapter.Fill(dataSet);
                EMPLmetroGrid1.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                EnableOff();
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile17_Click(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT TOP 1 Сотрудник.Фамилия , Сотрудник.Имя , Сотрудник.Отчество, Сотрудник.Должность,Сотрудник.Телефон,Сотрудник.Дата,Сотрудник.Пароль, Count(Продажа_абонемента.Идпродажа) AS [Количество продаж] ,Сотрудник.фото
                                                                         FROM Сотрудник 
                                                                         INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                         ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник 
                                                                         GROUP BY Сотрудник.Фамилия , Сотрудник.Имя , Сотрудник.Отчество, Сотрудник.Должность,Сотрудник.Телефон,Сотрудник.Дата,Сотрудник.Пароль, Сотрудник.Идсотрудник,сотрудник.фото  
                                                                         ORDER BY Count(Продажа_абонемента.Идпродажа) desc;",
                    selectConnection);
                adapter.Fill(dataSet);
                EMPLmetroGrid1.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                EnableOff();
                EMPLmetroTabControl3.Enabled = false;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile19_Click(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT Абонемент.Название AS Абонемент, Сотрудник.Фамилия  & Сотрудник.Имя AS [Реализовавший сотрудник], Спортсмен.Фамилия  & Спортсмен.Имя AS Спортсмен, Продажа_абонемента.Дата_начала as Начало,  Продажа_абонемента.Дата_окончания as Конец,сотрудник.фото,сотрудник.фото,сотрудник.фото,сотрудник.фото 
                              FROM Спортсмен 
                              INNER JOIN (Сотрудник 
                              INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                              ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник) 
                              ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;", selectConnection);
                adapter.Fill(dataSet);
                EMPLmetroGrid1.DataSource = dataSet.Tables[0];
                EMPLmetroGrid1.Columns[8].Visible = false;
                EMPLmetroGrid1.Columns[7].Visible = false;
                EMPLmetroGrid1.Columns[6].Visible = false;
                EMPLmetroGrid1.Columns[5].Visible = false;
                selectConnection.Close();
                EnableOff();
                EMPLmetroTabControl3.Enabled = false;
                if (EMPLmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile7_Click(object sender, EventArgs e)
        {
            var money = new Money();
            money.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            try
            {
                EMPLmetroTile3.Visible = true;
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                if (Convert.ToInt32(EMPLmetroTextBox6.Text) < Convert.ToInt32(EMPLmetroTextBox7.Text))
                {
                    var adapter = new OleDbDataAdapter(
                        $@"SELECT Сотрудник.Фамилия, Сотрудник.Имя, Сотрудник.Отчество, Сотрудник.Должность, Сотрудник.Телефон, Сотрудник.Дата, Сотрудник.Пароль, Зарплата_сотрудника.Зарплата ,Сотрудник.Фото
                                 FROM Зарплата_сотрудника 
                                 INNER JOIN Сотрудник 
                                 ON Зарплата_сотрудника.Идзарплата = Сотрудник.Идзарплата 
                                 where зарплата_сотрудника.зарплата 
                                 between {EMPLmetroTextBox6.Text} and {EMPLmetroTextBox7.Text}", selectConnection);
                    adapter.Fill(dataSet);
                    EMPLmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    EMPLmetroTile10.Enabled = false;
                    EMPLmetroTile11.Enabled = false;
                    EMPLmetroTabControl5.Enabled = false;
                    EMPLmetroTabControl7.Enabled = false;
                    EMPLmetroTabControl8.Enabled = false;
                    if (EMPLmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Сотрудника не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdEmployee();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальная зарплата не может быть больше конечной",
                        TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile21_Click(object sender, EventArgs e)
        {
            var sendComment = new SendComment();
            sendComment.ShowDialog();
        }

        private void metroTile23_Click(object sender, EventArgs e)
        {
            var reportEmployee = new ReportEmployee();
            reportEmployee.Show();
        }

        private void metroTile41_Click(object sender, EventArgs e)
        {
            var objSportsmenAdd = new ModSportsmen
            {
                textBox1 = {Text = ""},
                textBox2 = {Text = ""},
                textBox3 = {Text = ""},
                maskedTextBox1 = {Text = ""},
                metroDateTime1 = {Text = null},
                Text = @"Добавить спортсмена",
                metroTile1 = {Text = @"Добавить"}
            };
            if (objSportsmenAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    SPORTMmetroGrid2.Sort(SPORTMmetroGrid2.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddSportsmen = new OleDbCommand(@"INSERT INTO [спортсмен]
                                                        ( Фамилия, Имя, Отчество, Телефон, Дата_рождения, Пол)
                                                        VALUES(@st1,@st2,@st3,@st4,@st5,@st6)", connection);
                    queryAddSportsmen.Parameters.AddWithValue("st1", objSportsmenAdd.textBox1.Text);
                    queryAddSportsmen.Parameters.AddWithValue("st2", objSportsmenAdd.textBox2.Text);
                    queryAddSportsmen.Parameters.AddWithValue("st3", objSportsmenAdd.textBox3.Text);
                    queryAddSportsmen.Parameters.AddWithValue("st4", objSportsmenAdd.maskedTextBox1.Text);
                    queryAddSportsmen.Parameters.AddWithValue("st5",
                        Convert.ToDateTime(objSportsmenAdd.metroDateTime1.Text));
                    queryAddSportsmen.Parameters.AddWithValue("st6", objSportsmenAdd.metroComboBox1.Text);
                    queryAddSportsmen.ExecuteNonQuery();
                    connection.Close();
                    UpdSportsmen();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile40_Click(object sender, EventArgs e)
        {
            var objSportsmenUpdate = new ModSportsmen
            {
                Text = @"Редактировать спортсмена", metroTile1 = {Text = @"Редактировать"}
            };
            Debug.Assert(SPORTMmetroGrid2.CurrentRow != null, "Таблица пуста");
            idSportsmen = Convert.ToInt32(SPORTMmetroGrid2.CurrentRow.Cells[0].Value);
            objSportsmenUpdate.textBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[1].Value);
            objSportsmenUpdate.textBox2.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[2].Value);
            objSportsmenUpdate.textBox3.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[3].Value);
            objSportsmenUpdate.maskedTextBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[4].Value);
            objSportsmenUpdate.metroDateTime1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[5].Value);
            objSportsmenUpdate.metroComboBox1.Text = Convert.ToString(SPORTMmetroGrid2.CurrentRow.Cells[6].Value);
            if (objSportsmenUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    SPORTMmetroGrid2.Sort(SPORTMmetroGrid2.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateSportsmen =
                        new OleDbCommand(
                            "update спортсмен set Фамилия=@surname, Имя=@name, Отчество=@patryon, Телефон=@telephone, Дата_рождения=@date,Пол=@sex where идспортсмен=" +
                            idSportsmen + "", connection);
                    queryUpdateSportsmen.Parameters.AddWithValue("surname", objSportsmenUpdate.textBox1.Text);
                    queryUpdateSportsmen.Parameters.AddWithValue("name", objSportsmenUpdate.textBox2.Text);
                    queryUpdateSportsmen.Parameters.AddWithValue("patryon", objSportsmenUpdate.textBox3.Text);
                    queryUpdateSportsmen.Parameters.AddWithValue("telephone", objSportsmenUpdate.maskedTextBox1.Text);
                    queryUpdateSportsmen.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objSportsmenUpdate.metroDateTime1.Text));
                    queryUpdateSportsmen.Parameters.AddWithValue("sex", objSportsmenUpdate.metroComboBox1.Text);
                    queryUpdateSportsmen.ExecuteNonQuery();
                    connection.Close();
                    UpdSportsmen();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile39_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    connection.Open();
                    Debug.Assert(SPORTMmetroGrid2.CurrentRow != null, "Таблица пуста");
                    idSportsmen = Convert.ToInt32(SPORTMmetroGrid2.CurrentRow.Cells[0].Value);
                    var queryDeleteSportsmen = new OleDbCommand(@"DELETE FROM спортсмен 
                                                    WHERE идспортсмен=" + idSportsmen + "", connection);
                    queryDeleteSportsmen.ExecuteNonQuery();
                    UpdSportsmen();
                    connection.Close();
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile37_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.UpdEmployee();
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Sportsmen.docx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet) ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Спортсмен";
            excelApp.Cells[1, 1] = "Фамилия";
            excelApp.Cells[1, 2] = "Имя";
            excelApp.Cells[1, 3] = "Отчество";
            excelApp.Cells[1, 4] = "Телефон";
            excelApp.Cells[1, 5] = "Дата";
            excelApp.Cells[1, 6] = "Пол";
            {
                for (int i = 1; i < SPORTMmetroGrid2.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = SPORTMmetroGrid2.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = SPORTMmetroGrid2.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = SPORTMmetroGrid2.Rows[i - 1].Cells[3].Value;
                    excelApp.Cells[i + 1, 4] = SPORTMmetroGrid2.Rows[i - 1].Cells[4].Value;
                    excelApp.Cells[i + 1, 5] = SPORTMmetroGrid2.Rows[i - 1].Cells[5].Value;
                    excelApp.Cells[i + 1, 6] = SPORTMmetroGrid2.Rows[i - 1].Cells[6].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEmailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по спортсменам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEmailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по спортсменам", Body = "Отчет по спортсменам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587)
                            {
                                Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                EnableSsl = true
                            };
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                            objEmailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                        }
                }
            }
        }

        private void metroTile36_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Sportsmen.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application {Visible = true};
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                var table = doc.Tables.Add(range, SPORTMmetroGrid2.RowCount + 1, 6);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя";
                table.Cell(1, 3).Range.Text = "Отчество";
                table.Cell(1, 4).Range.Text = "Телефон";
                table.Cell(1, 5).Range.Text = "Дата";
                table.Cell(1, 6).Range.Text = "Пол";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по спортсменам", metroTile2 = {Text = @"Отправить"}
                        };

                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var objMessage = new MailMessage(myAddress, toAddress)
                                {
                                    Subject = "Отчет по спортсменам",
                                    Body = "Отчет по спортсменам",
                                    IsBodyHtml = true
                                };
                                objMessage.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587)
                                {
                                    Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                    EnableSsl = true
                                };
                                smtp.Send(objMessage);
                                MessageBox.Show(@"Сообщение отправлено на почту" + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                    }
                                }
                            }
                    }
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void HeadForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            var exit = new Exit();
            if (exit.ShowDialog() == DialogResult.OK)
            {
                Environment.Exit(0);
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
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                     MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox10_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl10.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                metroTile33.Enabled = false;
                metroTile32.Enabled = false;
                string queryFindSurnameSportsmen = @"SELECT *
                       FROM спортсмен 
                       WHERE[Фамилия] LIKE '%" + SPORTMtextBox10.Text + "%'";
                dataAdapterSportsmen = new OleDbDataAdapter(queryFindSurnameSportsmen, connection);
                dataTableSportsmen = new DataTable();
                dataAdapterSportsmen.Fill(dataTableSportsmen);
                SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SPORTMtextBox10.Text = "";
                    UpdSportsmen();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }


        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            try
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl10.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                metroTile34.Enabled = false;
                metroTile32.Enabled = false;
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                string date1 = SPORTMmetroDateTime3.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                string date2 = SPORTMmetroDateTime4.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                if (Convert.ToDateTime(SPORTMmetroDateTime3.Text) > Convert.ToDateTime(SPORTMmetroDateTime4.Text))
                {
                    var adapter = new OleDbDataAdapter($@"SELECT *
                                                                FROM спортсмен
                                                                WHERE Дата_рождения Between #{date1}# and #{date2}#",
                        selectConnection);
                    adapter.Fill(dataSet);
                    SPORTMmetroGrid2.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSportsmen();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdSportsmen();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (SPORTMmetroComboBox1.SelectedIndex)
            {
                case 0:
                    try
                    {
                        SPORTMmetroTabControl12.Enabled = false;
                        SPORTMmetroTabControl10.Enabled = false;
                        SPORTMmetroTabControl9.Enabled = false;
                        metroTile34.Enabled = false;
                        metroTile33.Enabled = false;
                        string selectSportsmenMen = @"SELECT * FROM Спортсмен WHERE (((Спортсмен.Пол) = 'Мужской'));";
                        dataAdapterSportsmen = new OleDbDataAdapter(selectSportsmenMen, connection);
                        dataTableSportsmen = new DataTable();
                        dataAdapterSportsmen.Fill(dataTableSportsmen);
                        SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                        if (SPORTMmetroGrid2.RowCount == 0)
                        {
                            MetroMessageBox.Show(this, "\nЗапись не найдена", "Спортсмена не найдено",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            UpdSportsmen();
                        }
                    }
                    catch (Exception exception)
                    {
                        MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        if (File.Exists(_fileNameLog) != true)
                        {
                            using (var streamWriter =
                                new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                            {
                                streamWriter.WriteLine(_dateLog);
                                streamWriter.WriteLine(exception.Message);
                            }
                        }
                        else
                        {
                            using (var streamWriter =
                                new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                            {
                                (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                streamWriter.WriteLine(_dateLog);
                                streamWriter.WriteLine(exception.Message);
                            }
                        }
                    }

                    break;
                case 1:
                    try
                    {
                        SPORTMmetroTabControl12.Enabled = false;
                        SPORTMmetroTabControl10.Enabled = false;
                        SPORTMmetroTabControl9.Enabled = false;
                        metroTile34.Enabled = false;
                        metroTile33.Enabled = false;
                        string selectSportesmenWomen =
                            @"SELECT * FROM Спортсмен WHERE (((Спортсмен.Пол) = 'Женский'));";
                        dataAdapterSportsmen = new OleDbDataAdapter(selectSportesmenWomen, connection);
                        dataTableSportsmen = new DataTable();
                        dataAdapterSportsmen.Fill(dataTableSportsmen);
                        SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                        if (SPORTMmetroGrid2.RowCount == 0)
                        {
                            MetroMessageBox.Show(this, "\nЗапись не найдена",
                                "Спортсменов с таким полом не найдено", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);

                            UpdSportsmen();
                        }
                    }
                    catch (Exception exception)
                    {
                        MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        if (File.Exists(_fileNameLog) != true)
                        {
                            using (var streamWriter =
                                new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                            {
                                streamWriter.WriteLine(_dateLog);
                                streamWriter.WriteLine(exception.Message);
                            }
                        }
                        else
                        {
                            using (var streamWriter =
                                new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                            {
                                (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                streamWriter.WriteLine(_dateLog);
                                streamWriter.WriteLine(exception.Message);
                            }
                        }
                    }

                    break;
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                UpdSportsmen();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroComboBox1_Click(object sender, EventArgs e)
        {
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            if (SPORTMtextBox9.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните пустые поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SPORTMmetroTabControl12.Enabled = false;
                    SPORTMmetroTabControl11.Enabled = false;
                    SPORTMmetroTabControl9.Enabled = false;
                    metroTile30.Enabled = false;
                    metroTile29.Enabled = false;
                    string selectFromСпортсменWhereФамилия = @"SELECT *
                           FROM спортсмен 
                           WHERE Фамилия='" + SPORTMtextBox9.Text + "'";
                    dataAdapterSportsmen = new OleDbDataAdapter(selectFromСпортсменWhereФамилия, connection);
                    dataTableSportsmen = new DataTable();
                    dataAdapterSportsmen.Fill(dataTableSportsmen);
                    SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                    SPORTMmetroGrid2.Columns[0].Visible = false;
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких спортсменов не найдено",
                            "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSportsmen();
                    }

                    SPORTMtextBox9.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {
            if (SPORTMtextBox7.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните пустые поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SPORTMmetroTabControl12.Enabled = false;
                    SPORTMmetroTabControl11.Enabled = false;
                    SPORTMmetroTabControl9.Enabled = false;
                    metroTile31.Enabled = false;
                    metroTile29.Enabled = false;
                    string selectFromСпортсменWhereИмя = @"SELECT *
                           FROM спортсмен 
                           WHERE Имя='" + SPORTMtextBox7.Text + "'";
                    dataAdapterSportsmen = new OleDbDataAdapter(selectFromСпортсменWhereИмя, connection);
                    dataTableSportsmen = new DataTable();
                    dataAdapterSportsmen.Fill(dataTableSportsmen);
                    SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                    SPORTMmetroGrid2.Columns[0].Visible = false;
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSportsmen();
                    }

                    SPORTMtextBox7.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            if ((SPORTMtextBox6.Text == "") || (SPORTMtextBox1.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе введены данные", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SPORTMmetroTabControl12.Enabled = false;
                    SPORTMmetroTabControl11.Enabled = false;
                    SPORTMmetroTabControl9.Enabled = false;
                    metroTile31.Enabled = false;
                    metroTile20.Enabled = false;
                    string selectFromСпортсменWhereФамилия = @"SELECT *
                           from спортсмен
                           WHERE Фамилия='" + SPORTMtextBox6.Text + "' and Имя='" + SPORTMtextBox1.Text + "'";
                    dataAdapterSportsmen = new OleDbDataAdapter(selectFromСпортсменWhereФамилия, connection);
                    dataTableSportsmen = new DataTable();
                    dataAdapterSportsmen.Fill(dataTableSportsmen);
                    SPORTMmetroGrid2.DataSource = dataTableSportsmen;
                    SPORTMmetroGrid2.Columns[0].Visible = false;
                    if (SPORTMmetroGrid2.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSportsmen();
                    }

                    SPORTMtextBox6.Text = "";
                    SPORTMtextBox1.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile28_Click(object sender, EventArgs e)
        {
            try
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl11.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                SPORTMmetroTabControl10.Enabled = false;
                SPORTMmetroTabControl14.Enabled = false;
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @" SELECT Спортсмен.Фамилия,Спортсмен.Имя,Спортсмен.Отчество, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Абонемент.Название, Абонемент.Количество_посещений
                                                            FROM Абонемент 
                                                            INNER JOIN (Спортсмен
                                                            INNER JOIN Продажа_абонемента
                                                            ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) 
                                                            ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент;",
                    selectConnection);
                adapter.Fill(dataSet);
                SPORTMmetroGrid2.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    UpdSportsmen();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile27_Click(object sender, EventArgs e)
        {
            SPORTMmetroTabControl12.Enabled = false;
            SPORTMmetroTabControl11.Enabled = false;
            SPORTMmetroTabControl9.Enabled = false;
            SPORTMmetroTabControl10.Enabled = false;
            SPORTMmetroTabControl14.Enabled = false;
            var selectConnection = new OleDbConnection(conString);
            selectConnection.Open();
            var dataSet = new DataSet();
            try
            {
                string inputBox = Microsoft.VisualBasic.Interaction.InputBox("Введите Фамилию спортсмена: ", "Запрос",
                    "", 500, 500);
                var adapter = new OleDbDataAdapter(String.Format(
                    @" SELECT Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество, Учет_посещений.Дата,Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания
FROM (Спортсмен INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) INNER JOIN Учет_посещений ON Продажа_абонемента.Идпродажа = Учет_посещений.Идпродажа
WHERE (((Спортсмен.Фамилия)='" + inputBox + "'));"), selectConnection);
                adapter.Fill(dataSet);
                SPORTMmetroGrid2.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    UpdSportsmen();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile26_Click(object sender, EventArgs e)
        {
            try
            {
                SPORTMmetroTabControl12.Enabled = false;
                SPORTMmetroTabControl11.Enabled = false;
                SPORTMmetroTabControl9.Enabled = false;
                SPORTMmetroTabControl10.Enabled = false;
                SPORTMmetroTabControl14.Enabled = false;
                var dbConnection = new OleDbConnection(conString);
                dbConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(@" SELECT Идспортсмен, Фамилия, Имя, Отчество,
       (SELECT COUNT(*) FROM Учет_посещений WHERE Идпродажа IN 
        (SELECT Идпродажа FROM продажа_абонемента WHERE Идспортсмен = Спортсмен.Идспортсмен)) AS [Количество посещений]
               FROM Спортсмен", dbConnection);
                adapter.Fill(dataSet);
                SPORTMmetroGrid2.DataSource = dataSet.Tables[0];
                dbConnection.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    UpdSportsmen();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile19_Click_1(object sender, EventArgs e)
        {
            var objTrenerAdd = new ModTrener
            {

                textBox1 = {Text = ""},
                textBox2 = {Text = ""},
                textBox3 = {Text = ""},
                maskedTextBox1 = {Text = ""},
                metroDateTime1 = {Text = null},
                metroTextBox5 = {Text = ""},
                label1= { Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[0].Value)},
            Text = @"Добавить тренера",
                metroTile1 = {Text = @"Добавить"},
                metroTextBox1 = {Text = ""}
            };
            if (objTrenerAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENmetroGrid1.Sort(TRENmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddTrener = new OleDbCommand(@"INSERT INTO [тренер]
                                                        ( Фамилия, Имя,          Отчество, Должность,        Телефон,       Дата_рождения,Пароль,    Фото,     оклад)
                                                    VALUES(@surname,  @name,      @patryon,     @position,    @telephone,       @date,     @pass,      @photo, @money)",
                        connection);
                    queryAddTrener.Parameters.AddWithValue("surname", objTrenerAdd.textBox1.Text);
                    queryAddTrener.Parameters.AddWithValue("name", objTrenerAdd.textBox2.Text);
                    queryAddTrener.Parameters.AddWithValue("patryon", objTrenerAdd.textBox3.Text);
                    queryAddTrener.Parameters.AddWithValue("position", objTrenerAdd.metroTextBox4.Text);
                    queryAddTrener.Parameters.AddWithValue("telephone", objTrenerAdd.maskedTextBox1.Text);
                    queryAddTrener.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objTrenerAdd.metroDateTime1.Text));
                    queryAddTrener.Parameters.AddWithValue("pass", Convert.ToInt32(objTrenerAdd.metroTextBox5.Text));
                    queryAddTrener.Parameters.AddWithValue("photo", objTrenerAdd.metroTextBox6.Text);
                    queryAddTrener.Parameters.AddWithValue("money", Convert.ToInt32(objTrenerAdd.metroTextBox1.Text));
                    queryAddTrener.ExecuteNonQuery();
                    connection.Close();
                    UpdTrener();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile18_Click_1(object sender, EventArgs e)
        {
            var objTrenerUpdate = new ModTrener
                {Text = @"Редактировать тренера", metroTile1 = {Text = @"Редактировать"}};
            Debug.Assert(TRENmetroGrid1.CurrentRow != null, "Таблица пуста");
            idTrener = Convert.ToInt32(TRENmetroGrid1.CurrentRow.Cells[0].Value);
            objTrenerUpdate.label1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[0].Value);
            objTrenerUpdate.textBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[1].Value);
            objTrenerUpdate.textBox2.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[2].Value);
            objTrenerUpdate.textBox3.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[3].Value);
            objTrenerUpdate.metroTextBox4.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[4].Value);
            objTrenerUpdate.maskedTextBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[5].Value);
            objTrenerUpdate.metroDateTime1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[6].Value);
            objTrenerUpdate.metroTextBox5.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[7].Value);
            objTrenerUpdate.metroTextBox6.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[8].Value);
            objTrenerUpdate.metroTextBox1.Text = Convert.ToString(TRENmetroGrid1.CurrentRow.Cells[9].Value);
            if (objTrenerUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENmetroGrid1.Sort(TRENmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateTrener =
                        new OleDbCommand(
                            "update тренер set Фамилия=@surname, Имя=@name, Отчество=@patryon, Должность=@position,Телефон=@phone, Дата_рождения=@date,Пароль=@pass,Фото=@photo,оклад=@money where идтренер=" +
                            idTrener + "", connection);
                    queryUpdateTrener.Parameters.AddWithValue("surname", objTrenerUpdate.textBox1.Text);
                    queryUpdateTrener.Parameters.AddWithValue("name", objTrenerUpdate.textBox2.Text);
                    queryUpdateTrener.Parameters.AddWithValue("patryon", objTrenerUpdate.textBox3.Text);
                    queryUpdateTrener.Parameters.AddWithValue("position", objTrenerUpdate.metroTextBox4.Text);
                    queryUpdateTrener.Parameters.AddWithValue("phone", objTrenerUpdate.maskedTextBox1.Text);
                    queryUpdateTrener.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objTrenerUpdate.metroDateTime1.Text));
                    queryUpdateTrener.Parameters.AddWithValue("pass",
                        Convert.ToInt32(objTrenerUpdate.metroTextBox5.Text));
                    queryUpdateTrener.Parameters.AddWithValue("photo", objTrenerUpdate.metroTextBox6.Text);
                    queryUpdateTrener.Parameters.AddWithValue("money",
                        Convert.ToInt32(objTrenerUpdate.metroTextBox1.Text));
                    queryUpdateTrener.ExecuteNonQuery();
                    connection.Close();
                    UpdTrener();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile17_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                connection.Open();
                Debug.Assert(TRENmetroGrid1.CurrentRow != null, "Таблица пуста");
                idTrener = Convert.ToInt32(TRENmetroGrid1.CurrentRow.Cells[0].Value);
                var queryDeleteTrener = new OleDbCommand(@"DELETE FROM тренер 
                                                    WHERE идтренер=" + idTrener + "", connection);
                queryDeleteTrener.ExecuteNonQuery();
                UpdTrener();
                connection.Close();
            }
        }

        private void metroTile15_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.UpdEmployee();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet) ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Тренер";
            excelApp.Cells[1, 1] = "Фамилия";
            excelApp.Cells[1, 2] = "Имя";
            excelApp.Cells[1, 3] = "Отчество";
            excelApp.Cells[1, 4] = "Должность";
            excelApp.Cells[1, 5] = "Телефон";
            excelApp.Cells[1, 6] = "Дата";
            excelApp.Cells[1, 7] = "Пароль";
            excelApp.Cells[1, 8] = "Оклад";
            {
                for (int i = 1; i < TRENmetroGrid1.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = TRENmetroGrid1.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = TRENmetroGrid1.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = TRENmetroGrid1.Rows[i - 1].Cells[3].Value;
                    excelApp.Cells[i + 1, 4] = TRENmetroGrid1.Rows[i - 1].Cells[4].Value;
                    excelApp.Cells[i + 1, 5] = TRENmetroGrid1.Rows[i - 1].Cells[5].Value;
                    excelApp.Cells[i + 1, 6] = TRENmetroGrid1.Rows[i - 1].Cells[6].Value;
                    excelApp.Cells[i + 1, 7] = TRENmetroGrid1.Rows[i - 1].Cells[7].Value;
                    excelApp.Cells[i + 1, 8] = TRENmetroGrid1.Rows[i - 1].Cells[9].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                string path = Directory.GetCurrentDirectory() + @"\" + "report/Trener.docx";
                if (File.Exists(path))
                {
                    File.Delete(path);
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEMailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по тренерам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEMailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEMailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по тренерам", Body = "Отчет по тренерам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587);
                            smtp.Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq");
                            smtp.EnableSsl = true;
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту" + objEMailReport.metroTextBox4.Text);
                            objEMailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                "Сообщение не отправлено на почту " + objEMailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                        }
                }
            }
        }

        private void metroTile14_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trener.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application {Visible = true};
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, TRENmetroGrid1.RowCount + 1, 8);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Фамилия";
                table.Cell(1, 2).Range.Text = "Имя";
                table.Cell(1, 3).Range.Text = "Отчество";
                table.Cell(1, 4).Range.Text = "Должность";
                table.Cell(1, 5).Range.Text = "Телефон";
                table.Cell(1, 6).Range.Text = "Дата";
                table.Cell(1, 7).Range.Text = "Пароль";
                table.Cell(1, 8).Range.Text = "Оклад";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по тренерам", metroTile2 = {Text = @"Отправить"}
                        };
                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var objMessage = new MailMessage(myAddress, toAddress)
                                {
                                    Subject = "Отчет по тренерам", Body = "Отчет по тренерам", IsBodyHtml = true
                                };
                                objMessage.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587);
                                smtp.Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq");
                                smtp.EnableSsl = true;
                                smtp.Send(objMessage);
                                MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this,
                                    "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                    "Не отправлено",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                FocusMe();
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                            }
                    }
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                     MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox8_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                string selectTrenerwhereSurname =
                    @"SELECT Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                       FROM тренер
                       WHERE[Фамилия] LIKE '%" + TRENtextBox8.Text + "%'";
                dataAdapterTrener = new OleDbDataAdapter(selectTrenerwhereSurname, connection);
                dataTableTrener = new DataTable();
                dataAdapterTrener.Fill(dataTableTrener);
                TRENmetroGrid1.DataSource = dataTableTrener;
                metroTile11.Enabled = false;
                metroTile10.Enabled = false;
                TRENmetroTabControl4.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    TRENtextBox8.Text = "";
                    UpdTrener();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton5_Click_1(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                string date1 = TRENmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                string date2 = TRENmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                if (Convert.ToDateTime(TRENmetroDateTime2.Text) < Convert.ToDateTime(TRENmetroDateTime1.Text))
                {
                    var adapter = new OleDbDataAdapter
                    ($@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                                WHERE Дата_рождения Between #{date1}# and #{date2}#", selectConnection);
                    adapter.Fill(dataSet);
                    TRENmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    metroTile12.Enabled = false;
                    metroTile10.Enabled = false;
                    TRENmetroTabControl4.Enabled = false;
                    TRENmetroTabControl3.Enabled = false;
                    TRENmetroTabControl6.Enabled = false;
                    if (TRENmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrener();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdTrener();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton4_Click_1(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                if (Convert.ToInt32(TRENmetroTextBox3.Text) < Convert.ToInt32(TRENmetroTextBox2.Text))
                {
                    var adapter = new OleDbDataAdapter(
                        $@"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер 
                               where оклад between {TRENmetroTextBox3.Text} and {TRENmetroTextBox2.Text}",
                        selectConnection);
                    adapter.Fill(dataSet);
                    TRENmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    metroTile12.Enabled = false;
                    metroTile11.Enabled = false;
                    TRENmetroTabControl4.Enabled = false;
                    TRENmetroTabControl3.Enabled = false;
                    TRENmetroTabControl6.Enabled = false;
                    if (TRENmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренера не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrener();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальный оклад не может быть больше конечного",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdTrener();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton3_Click_1(object sender, EventArgs e)
        {
            if (TRENtextBox5.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string selectTrenerWhereSurname =
                        @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Фамилия='" + TRENtextBox5.Text + "'";
                    dataAdapterTrener = new OleDbDataAdapter(selectTrenerWhereSurname, connection);
                    dataTableTrener = new DataTable();
                    dataAdapterTrener.Fill(dataTableTrener);
                    TRENmetroGrid1.DataSource = dataTableTrener;
                    TRENmetroGrid1.Columns[8].Visible = false;
                    metroTile7.Enabled = false;
                    metroTile8.Enabled = false;
                    TRENmetroTabControl5.Enabled = false;
                    TRENmetroTabControl3.Enabled = false;
                    TRENmetroTabControl6.Enabled = false;
                    if (TRENmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrener();
                    }

                    TRENtextBox5.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton2_Click_1(object sender, EventArgs e)
        {
            if (TRENtextBox4.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string s = @"select Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                           WHERE Имя='" + TRENtextBox4.Text + "'";
                    dataAdapterTrener = new OleDbDataAdapter(s, connection);
                    dataTableTrener = new DataTable();
                    dataAdapterTrener.Fill(dataTableTrener);
                    TRENmetroGrid1.DataSource = dataTableTrener;
                    TRENmetroGrid1.Columns[8].Visible = false;
                    metroTile7.Enabled = false;
                    metroTile9.Enabled = false;
                    TRENmetroTabControl5.Enabled = false;
                    TRENmetroTabControl3.Enabled = false;
                    TRENmetroTabControl6.Enabled = false;
                    if (TRENmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrener();
                    }

                    TRENtextBox4.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            if ((TRENtextBox3.Text == "") || (TRENtextBox2.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе введены данные", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string selectTrenerwhereSurname =
                        @"SELECT Фамилия, Имя, Отчество, Должность, Телефон, Дата_рождения, Пароль,оклад, Фото
                               FROM тренер
                               WHERE Фамилия='" + TRENtextBox3.Text + "' and Имя='" + TRENtextBox2.Text + "'";
                    dataAdapterTrener = new OleDbDataAdapter(selectTrenerwhereSurname, connection);
                    dataTableTrener = new DataTable();
                    dataAdapterTrener.Fill(dataTableTrener);
                    TRENmetroGrid1.DataSource = dataTableTrener;
                    TRENmetroGrid1.Columns[8].Visible = false;
                    metroTile8.Enabled = false;
                    metroTile9.Enabled = false;
                    TRENmetroTabControl5.Enabled = false;
                    TRENmetroTabControl3.Enabled = false;
                    TRENmetroTabControl6.Enabled = false;
                    if (TRENmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренера не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrener();
                    }

                    TRENtextBox3.Text = "";
                    TRENtextBox2.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
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
                MetroMessageBox.Show(this, ex.Message, "Ошибка");
            }
        }

        private void сброситьФильтрToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                UpdTrener();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile6_Click_1(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT Тренер.Фамилия, Тренер.Имя, Тренер.Отчество, Спортсмен.Фамилия, Спортсмен.Имя, Спортсмен.Отчество,Тренировка.Название,тренер.Телефон,Тренер.Фото
                                                  FROM Спортсмен 
                                                   INNER JOIN (((Тренер INNER JOIN Тренировка ON Тренер.Идтренер = Тренировка.Идтренер)
                                                    INNER JOIN Абонемент ON Тренировка.Идтренировка = Абонемент.Идтренеровка) 
                                                    INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                    ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;",
                    selectConnection);
                metroTile16.Visible = false;
                TRENpictureBox2.Visible = false;
                adapter.Fill(dataSet);
                TRENmetroGrid1.DataSource = dataSet.Tables[0];
                TRENmetroTabControl8.Enabled = false;
                TRENmetroTabControl6.Enabled = false;
                TRENmetroTabControl3.Enabled = false;
                TRENmetroTabControl4.Enabled = false;
                TRENmetroTabControl5.Enabled = false;
                selectConnection.Close();
                TRENmetroGrid1.Columns[8].Visible = false;
                if (TRENmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdTrener();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile13_Click(object sender, EventArgs e)
        {
            var reportTrener = new ReportTrener();
            reportTrener.Show();
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            var reportSportsmen = new ReportSportsmen();
            reportSportsmen.Show();
        }

        private void metroTile38_Click(object sender, EventArgs e)
        {
            var objTrenerovkaAdd = new ModTrenerovka
            {
                textBox1 = {Text = ""}, Text = @"Добавить тренировку", metroTile1 = {Text = @"Добавить"}
            };
            connection.Open();
            var cmd =
                new OleDbCommand(@"SELECT Вид_тренировки.Идвидтренировка, Вид_тренировки.Название FROM Вид_тренировки;",
                    connection);
            objTrenerovkaAdd.metroComboBox1.DisplayMember = "Название";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objTrenerovkaAdd.metroComboBox1.DataSource = list.ToList();
            objTrenerovkaAdd.metroComboBox1.DisplayMember = "Value";
            objTrenerovkaAdd.metroComboBox1.ValueMember = "Key";
            connection.Close();
            connection.Open();
            var cmd1 = new OleDbCommand(@"SELECT Тренер.Идтренер, Тренер.Фамилия FROM Тренер;", connection);
            objTrenerovkaAdd.metroComboBox2.DisplayMember = "Фамилия";
            var reader1 = cmd1.ExecuteReader();
            var list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int) reader1[0], (string) reader1[1]);
            }

            reader1.Close();
            cmd1.ExecuteNonQuery();
            objTrenerovkaAdd.metroComboBox2.DataSource = list1.ToList();
            objTrenerovkaAdd.metroComboBox2.DisplayMember = "Value";
            objTrenerovkaAdd.metroComboBox2.ValueMember = "Key";
            connection.Close();
            if (objTrenerovkaAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    TRENINGmetroGrid1.Sort(TRENINGmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddTrening = new OleDbCommand(@"INSERT INTO [тренировка]
                                                        ( Название, Идвидтренировка, Идтренер)
                                                        VALUES(@name,@idtrening,@idtrener)", connection);
                    queryAddTrening.Parameters.AddWithValue("name", objTrenerovkaAdd.textBox1.Text);
                    queryAddTrening.Parameters.AddWithValue("idtrening",
                        Convert.ToInt32(objTrenerovkaAdd.metroComboBox1.SelectedValue.ToString()));
                    queryAddTrening.Parameters.AddWithValue("idtrener",
                        Convert.ToInt32(objTrenerovkaAdd.metroComboBox2.SelectedValue.ToString()));
                    queryAddTrening.ExecuteNonQuery();
                    connection.Close();
                    UpdTrening();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile37_Click_1(object sender, EventArgs e)
        {
            var objTrenerovkaUpdate = new ModTrenerovka();
            connection.Close();
            objTrenerovkaUpdate.Text = @"Редактировать тренировку";
            objTrenerovkaUpdate.metroTile1.Text = @"Редактировать";
            Debug.Assert(TRENINGmetroGrid1.CurrentRow != null, "Таблица пуста");
            idTrenerovka = Convert.ToInt32(TRENINGmetroGrid1.CurrentRow.Cells[0].Value);
            objTrenerovkaUpdate.textBox1.Text = Convert.ToString(TRENINGmetroGrid1.CurrentRow.Cells[1].Value);
            connection.Open();
            var cmd =
                new OleDbCommand(@"SELECT Вид_тренировки.Идвидтренировка, Вид_тренировки.Название FROM Вид_тренировки;",
                    connection);
            objTrenerovkaUpdate.metroComboBox1.DisplayMember = "Название";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objTrenerovkaUpdate.metroComboBox1.DataSource = list.ToList();
            objTrenerovkaUpdate.metroComboBox1.DisplayMember = "Value";
            objTrenerovkaUpdate.metroComboBox1.ValueMember = "Key";
            objTrenerovkaUpdate.metroComboBox1.SelectedValue = TRENINGmetroGrid1.CurrentRow.Cells[5].Value;
            connection.Close();
            connection.Open();
            var cmd1 = new OleDbCommand(@"SELECT Тренер.Идтренер, Тренер.Фамилия FROM Тренер;", connection);
            objTrenerovkaUpdate.metroComboBox2.DisplayMember = "Фамилия";
            var reader1 = cmd1.ExecuteReader();
            var list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int) reader1[0], (string) reader1[1]);
            }

            reader1.Close();
            cmd1.ExecuteNonQuery();
            objTrenerovkaUpdate.metroComboBox2.DataSource = list1.ToList();
            objTrenerovkaUpdate.metroComboBox2.DisplayMember = "Value";
            objTrenerovkaUpdate.metroComboBox2.ValueMember = "Key";
            objTrenerovkaUpdate.metroComboBox2.SelectedValue = TRENINGmetroGrid1.CurrentRow.Cells[4].Value;
            connection.Close();
            if (objTrenerovkaUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    TRENINGmetroGrid1.Sort(TRENINGmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateTrening =
                        new OleDbCommand(
                            "update тренировка set название=@name, Идвидтренировка=@idtrening, Идтренер=@idtrener where Идтренировка=" +
                            idTrenerovka + "", connection);
                    queryUpdateTrening.Parameters.AddWithValue("name", objTrenerovkaUpdate.textBox1.Text);
                    queryUpdateTrening.Parameters.AddWithValue("idtrening",
                        Convert.ToInt32(objTrenerovkaUpdate.metroComboBox1.SelectedValue.ToString()));
                    queryUpdateTrening.Parameters.AddWithValue("idtrener",
                        Convert.ToInt32(objTrenerovkaUpdate.metroComboBox2.SelectedValue.ToString()));
                    queryUpdateTrening.ExecuteNonQuery();
                    connection.Close();
                    UpdTrening();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile36_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    connection.Close();
                    connection.Open();
                    Debug.Assert(TRENINGmetroGrid1.CurrentRow != null, "Таблица пуста");
                    idTrenerovka = Convert.ToInt32(TRENINGmetroGrid1.CurrentRow.Cells[0].Value);
                    var queryDeleteTrening = new OleDbCommand(@"DELETE FROM тренировка 
                                                    WHERE идтренировка=" + idTrenerovka + "", connection);
                    queryDeleteTrening.ExecuteNonQuery();
                    UpdTrening();
                    connection.Close();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile28_Click_1(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trenerovka.docx";
            this.UpdEmployee();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet) ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Сотрудники";
            excelApp.Cells[1, 1] = "Название";
            excelApp.Cells[1, 2] = "Вид";
            excelApp.Cells[1, 3] = "Тренер";
            {
                for (int i = 1; i < TRENINGmetroGrid1.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = TRENINGmetroGrid1.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = TRENINGmetroGrid1.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = TRENINGmetroGrid1.Rows[i - 1].Cells[3].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEmailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по тренировкам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEmailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по тренировкам", Body = "Отчет по тренировкам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587)
                            {
                                Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                EnableSsl = true
                            };
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                            objEmailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                        }
                }
            }
        }

        private void metroTile27_Click_1(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Trenerovka.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application();
            wordapp.Visible = true;
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                var table = doc.Tables.Add(range, TRENINGmetroGrid1.RowCount + 1, 3);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Название";
                table.Cell(1, 2).Range.Text = "Вид";
                table.Cell(1, 3).Range.Text = "Тренер";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < TRENINGmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = TRENINGmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                }

                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по тренировкам", metroTile2 = {Text = @"Отправить"}
                        };
                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var objMessage = new MailMessage(myAddress, toAddress);
                                objMessage.Subject = "Отчет по тренировкам";
                                objMessage.Body = "Отчет по тренировкам";
                                objMessage.IsBodyHtml = true;
                                objMessage.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587);
                                smtp.Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq");
                                smtp.EnableSsl = true;
                                smtp.Send(objMessage);
                                MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this,
                                    "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                    "Не отправлено",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                FocusMe();
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                            }
                    }
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void textBox6_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox5_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                     MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                string selectTreningWhereLike =
                    @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер 
                       WHERE[Тренировка.Название] LIKE '%" + TRENINGtextBox6.Text + "%'";
                dataAdapterTrening = new OleDbDataAdapter(selectTreningWhereLike, connection);
                dataTableTrenerovka = new DataTable();
                dataAdapterTrening.Fill(dataTableTrenerovka);
                TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
                TRENINGmetroTabControl6.Enabled = false;
                //metroTile18.Enabled = false;
                TRENINGmetroTabControl4.Enabled = false;
                TRENINGmetroTabControl3.Enabled = false;
                if (TRENINGmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Тренировки не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    EMPLtextBox2.Text = "";
                    UpdTrening();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton3_Click_2(object sender, EventArgs e)
        {
            if (TRENINGtextBox5.Text == "")
            {
                MetroMessageBox.Show(this, "\nЗаполните все поля", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    string selectViewTreningwhere =
                        @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Вид_тренировки.Название='" + TRENINGtextBox5.Text + "'";
                    dataAdapterTrening = new OleDbDataAdapter(selectViewTreningwhere, connection);
                    dataTableTrenerovka = new DataTable();
                    dataAdapterTrening.Fill(dataTableTrenerovka);
                    TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
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
                        MetroMessageBox.Show(this, "\nТаких видов тренировок не найдено",
                            "Вида тренировки не найдено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrening();
                    }

                    TRENINGtextBox5.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void сброситьФильтрToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            try
            {
                UpdTrening();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton2_Click_2(object sender, EventArgs e)
        {
            if (TRENINGtextBox4.Text == "")
            {
                MetroMessageBox.Show(this, "\nЗаполните все поля", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    string selectTrenerWhereSurname =
                        @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Тренер.Фамилия='" + TRENINGtextBox4.Text + "'";
                    dataAdapterTrening = new OleDbDataAdapter(selectTrenerWhereSurname, connection);
                    dataTableTrenerovka = new DataTable();
                    dataAdapterTrening.Fill(dataTableTrenerovka);
                    TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
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
                        MetroMessageBox.Show(this, "\nТаких тренеров не найдено", "Тренеров не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrening();
                    }

                    TRENINGtextBox4.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton1_Click_2(object sender, EventArgs e)
        {
            if (TRENINGtextBox3.Text == "")
            {
                MetroMessageBox.Show(this, "\nЗаполните все поля", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    string selectTreningwhereName =
                        @"SELECT Тренировка.Идтренировка, Тренировка.Название, Вид_тренировки.Название, Тренер.Фамилия,тренер.идтренер, вид_тренировки.идвидтренировка
FROM Тренер INNER JOIN (Вид_тренировки INNER JOIN Тренировка ON Вид_тренировки.Идвидтренировка = Тренировка.Идвидтренировка) ON Тренер.Идтренер = Тренировка.Идтренер
 WHERE Тренировка.Название='" + TRENINGtextBox3.Text + "'";
                    dataAdapterTrening = new OleDbDataAdapter(selectTreningwhereName, connection);
                    dataTableTrenerovka = new DataTable();
                    dataAdapterTrening.Fill(dataTableTrenerovka);
                    TRENINGmetroGrid1.DataSource = dataTableTrenerovka;
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
                        MetroMessageBox.Show(this, "\nТаких тренировок не найдено", "Тренировки не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdTrening();
                    }

                    TRENINGtextBox3.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroComboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
        }

        private void metroTile26_Click_1(object sender, EventArgs e)
        {
            var reportTrening = new ReportTrening();
            reportTrening.Show();
        }

        private void metroTile35_Click(object sender, EventArgs e)
        {
            var viewTrening = new ViewTrening();
            viewTrening.ShowDialog();
        }

        public void pictureBox2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Normal;
            Height = -9999;
            Width = -9999;
            }

        private void metroTile13_Click_1(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT top 1 Тренировка.Идтренировка, тренировка.название, count(Тренировка.Идтренировка) as [В абонементах]
            FROM Тренировка INNER JOIN Абонемент ON Тренировка.Идтренировка = Абонемент.Идтренеровка
            GROUP BY Тренировка.Идтренировка, Тренировка.Название
            ORDER BY sum(Тренировка.Идтренировка)desc", selectConnection);
                TRENINGmetroTabControl3.Enabled = false;
                TRENINGmetroTabControl5.Enabled = false;
                TRENINGmetroTabControl4.Enabled = false;
                TRENINGmetroTabControl6.Enabled = false;
                TRENINGmetroTabControl7.Enabled = false;
                TRENINGmetroTabControl8.Enabled = false;
                adapter.Fill(dataSet);
                TRENINGmetroGrid1.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                if (SPORTMmetroGrid2.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdTrening();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile41_Click_1(object sender, EventArgs e)
        {
            var objAbonementAdd = new ModAbonement
            {
                textBox1 = {Text = ""},
                textBox2 = {Text = ""},
                metroLabel4= { Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[0].Value)},
            textBox3 = {Text = ""},
                Text = @"Добавить абонемент",
                metroTile1 = {Text = @"Добавить"}
            };
            connection.Open();
            var cmd = new OleDbCommand("SELECT Тренировка.Идтренировка, Тренировка.Название FROM Тренировка",
                connection);
            objAbonementAdd.metroComboBox1.DisplayMember = "Название";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objAbonementAdd.metroComboBox1.DataSource = list.ToList();
            objAbonementAdd.metroComboBox1.DisplayMember = "Value";
            objAbonementAdd.metroComboBox1.ValueMember = "Key";
            connection.Close();
            if (objAbonementAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    ABONmetroGrid1.Sort(ABONmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddAbonement = new OleDbCommand(@"INSERT INTO [абонемент]
                                                        ( Название, Цена, Количество_посещений, Идтренеровка)
                                                        VALUES(@name,@money,@count,@idtrening)", connection);
                    queryAddAbonement.Parameters.AddWithValue("name", objAbonementAdd.textBox1.Text);
                    queryAddAbonement.Parameters.AddWithValue("money", objAbonementAdd.textBox2.Text);
                    queryAddAbonement.Parameters.AddWithValue("count", objAbonementAdd.textBox3.Text);
                    queryAddAbonement.Parameters.AddWithValue("idtrening",
                        Convert.ToInt32(objAbonementAdd.metroComboBox1.SelectedValue.ToString()));
                    queryAddAbonement.ExecuteNonQuery();
                    connection.Close();
                    UpdAbonement();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile40_Click_1(object sender, EventArgs e)
        {
            var objAbonementUpdate = new ModAbonement();
            connection.Close();
            objAbonementUpdate.Text = @"Редактировать абонемент";
            objAbonementUpdate.metroTile1.Text = @"Редактировать";
            Debug.Assert(ABONmetroGrid1.CurrentRow != null, "Таблица пуста");
            idAbonement = Convert.ToInt32(ABONmetroGrid1.CurrentRow.Cells[0].Value);
            objAbonementUpdate.metroLabel4.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[0].Value);
            objAbonementUpdate.textBox1.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[1].Value);
            objAbonementUpdate.textBox2.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[2].Value);
            objAbonementUpdate.textBox3.Text = Convert.ToString(ABONmetroGrid1.CurrentRow.Cells[3].Value);
            connection.Open();
            var cmd = new OleDbCommand("SELECT Тренировка.Идтренировка, Тренировка.Название FROM Тренировка",
                connection);
            objAbonementUpdate.metroComboBox1.DisplayMember = "Название";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objAbonementUpdate.metroComboBox1.DataSource = list.ToList();
            objAbonementUpdate.metroComboBox1.DisplayMember = "Value";
            objAbonementUpdate.metroComboBox1.ValueMember = "Key";
            connection.Close();
            objAbonementUpdate.metroComboBox1.SelectedValue = ABONmetroGrid1.CurrentRow.Cells[5].Value;
            if (objAbonementUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    ABONmetroGrid1.Sort(ABONmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateAbonement =
                        new OleDbCommand(
                            "update абонемент set Название=@name, Цена=@money, Количество_посещений=@count, Идтренеровка=@idtrening where идабонемент=" +
                            idAbonement + "", connection);
                    queryUpdateAbonement.Parameters.AddWithValue("name", objAbonementUpdate.textBox1.Text);
                    queryUpdateAbonement.Parameters.AddWithValue("money", objAbonementUpdate.textBox2.Text);
                    queryUpdateAbonement.Parameters.AddWithValue("count", objAbonementUpdate.textBox3.Text);
                    queryUpdateAbonement.Parameters.AddWithValue("idtrening",
                        Convert.ToInt32(objAbonementUpdate.metroComboBox1.SelectedValue));
                    queryUpdateAbonement.ExecuteNonQuery();
                    connection.Close();
                    UpdAbonement();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile39_Click_1(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    connection.Close();
                    connection.Open();
                    Debug.Assert(ABONmetroGrid1.CurrentRow != null, "Таблица пуста");
                    idAbonement = Convert.ToInt32(ABONmetroGrid1.CurrentRow.Cells[0].Value);
                    var queryDeleteAbonement = new OleDbCommand(@"DELETE FROM абонемент 
                                                    WHERE идабонемент=" + idAbonement + "", connection);
                    queryDeleteAbonement.ExecuteNonQuery();
                    UpdAbonement();
                    connection.Close();
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile38_Click_1(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Abonement.docx";
            this.UpdAbonement();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet) ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Абонемент";
            excelApp.Cells[1, 1] = "Название";
            excelApp.Cells[1, 2] = "Цена";
            excelApp.Cells[1, 3] = "Количество";
            excelApp.Cells[1, 4] = "Тренеровка";
            {
                for (int i = 1; i < ABONmetroGrid1.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = ABONmetroGrid1.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = ABONmetroGrid1.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = ABONmetroGrid1.Rows[i - 1].Cells[3].Value;
                    excelApp.Cells[i + 1, 4] = ABONmetroGrid1.Rows[i - 1].Cells[4].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEmailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по абонементам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEmailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по абонементам", Body = "Отчет по абонементам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587)
                            {
                                Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                EnableSsl = true
                            };
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                            objEmailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                        }
                }
            }
        }

        private void metroTile37_Click_2(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/Abonement.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application {Visible = true};
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                var table = doc.Tables.Add(range, ABONmetroGrid1.RowCount + 1, 4);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Название";
                table.Cell(1, 2).Range.Text = "Цена";
                table.Cell(1, 3).Range.Text = "Количество";
                table.Cell(1, 4).Range.Text = "Тренировка";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                for (int i = 1; i < ABONmetroGrid1.RowCount + 1; i++)
                {
                    table.Cell(i + 1, 1).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[1].Value.ToString();
                    table.Cell(i + 1, 2).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[2].Value.ToString();
                    table.Cell(i + 1, 3).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[3].Value.ToString();
                    table.Cell(i + 1, 4).Range.Text = ABONmetroGrid1.Rows[i - 1].Cells[4].Value.ToString();
                }

                try
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по абонементам", metroTile2 = {Text = @"Отправить"}
                        };
                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var objMessage = new MailMessage(myAddress, toAddress)
                                {
                                    Subject = "Отчет по абонементам",
                                    Body = "Отчет по абонементам",
                                    IsBodyHtml = true
                                };
                                objMessage.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587)
                                {
                                    Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                    EnableSsl = true
                                };
                                smtp.Send(objMessage);
                                MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this,
                                    "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                    "Не отправлено",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                FocusMe();
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                            }
                    }
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile36_Click_2(object sender, EventArgs e)
        {
            var reportSportsmen = new ReportSportsmen();
            reportSportsmen.Show();
        }

        private void сброситьФильтрToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            try
            {
                UpdAbonement();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton1_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox3.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните все поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string selectTreningWhereName =
                        @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE  Тренировка.Название='" + ABONtextBox3.Text + "'";
                    dataAdapterAbonement = new OleDbDataAdapter(selectTreningWhereName, connection);
                    dataTableAbonement = new DataTable();
                    dataAdapterAbonement.Fill(dataTableAbonement);
                    ABONmetroGrid1.DataSource = dataTableAbonement;
                    ABONmetroTabControl7.Enabled = false;
                    ABONmetroTabControl6.Enabled = false;
                    ABONmetroTabControl4.Enabled = false;
                    metroTile19.Enabled = false;
                    metroTile26.Enabled = false;
                    if (ABONmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких абонементов не найдено", "Абонемента не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdAbonement();
                    }

                    ABONtextBox3.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile6_Click_3(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT top 1  абонемент.идабонемент,абонемент.название, count (абонемент.идабонемент) as [В продажах]
FROM Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент
 GROUP BY  абонемент.идабонемент,абонемент.название
            ORDER BY sum(абонемент.идабонемент)desc", selectConnection);
                adapter.Fill(dataSet);
                ABONmetroGrid1.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                ABONmetroTabControl4.Enabled = false;
                ABONmetroTabControl5.Enabled = false;
                ABONmetroTabControl6.Enabled = false;
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl8.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdAbonement();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile28_Click_2(object sender, EventArgs e)
        {
            var reportAbonement = new ReportAbonement();
            reportAbonement.Show();
        }

        private void textBox8_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                     MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox3_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox8_TextChanged_1(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox3_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox8_KeyUp_1(object sender, KeyEventArgs e)
        {
            try
            {
                string selectAbonementWhereName =
                    @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                       WHERE[Абонемент.Название] LIKE '%" + ABONtextBox8.Text + "%'";
                dataAdapterAbonement = new OleDbDataAdapter(selectAbonementWhereName, connection);
                dataTableAbonement = new DataTable();
                dataAdapterAbonement.Fill(dataTableAbonement);
                ABONmetroGrid1.DataSource = dataTableAbonement;
                ABONmetroTabControl7.Enabled = false;
                ABONmetroTabControl5.Enabled = false;
                ABONmetroTabControl4.Enabled = false;
                metroTile27.Enabled = false;
                if (ABONmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемент не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ABONtextBox8.Text = "";
                    UpdAbonement();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton4_Click_2(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                if (Convert.ToInt32(ABONmetroTextBox3.Text) < Convert.ToInt32(ABONmetroTextBox2.Text))
                {
                    var adapter = new OleDbDataAdapter(
                        $@"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка  where Абонемент.Цена between {ABONmetroTextBox3.Text} and {ABONmetroTextBox2.Text}",
                        selectConnection);
                    adapter.Fill(dataSet);
                    ABONmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    ABONmetroTabControl7.Enabled = false;
                    ABONmetroTabControl5.Enabled = false;
                    ABONmetroTabControl4.Enabled = false;
                    metroTile35.Enabled = false;
                    if (ABONmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемента не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdAbonement();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальная цена не может быть больше конечной",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdAbonement();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton3_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox5.Text == "")
            {
                MetroMessageBox.Show(this, @"Пустые(ое) поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string selectWhereAbonementName =
                        @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE Абонемент.Название='" + ABONtextBox5.Text + "'";
                    dataAdapterAbonement = new OleDbDataAdapter(selectWhereAbonementName, connection);
                    dataTableAbonement = new DataTable();
                    dataAdapterAbonement.Fill(dataTableAbonement);
                    ABONmetroGrid1.DataSource = dataTableAbonement;
                    ABONmetroTabControl7.Enabled = false;
                    ABONmetroTabControl6.Enabled = false;
                    ABONmetroTabControl4.Enabled = false;
                    metroTile19.Enabled = false;
                    metroTile5.Enabled = false;
                    if (ABONmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких сотрудников не найдено", "Абонемента не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdAbonement();
                    }

                    ABONtextBox5.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton2_Click_3(object sender, EventArgs e)
        {
            if (ABONtextBox4.Text == "")
            {
                MetroMessageBox.Show(this, "Заполните все поля", Title, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string selectAbonementWhereCount =
                        @"SELECT Абонемент.Название, Абонемент.Цена, Абонемент.Количество_посещений, Тренировка.Название, Абонемент.Идтренеровка
                                                     FROM Тренировка INNER JOIN Абонемент
                                                     ON Тренировка.Идтренировка = Абонемент.Идтренеровка 
                           WHERE Абонемент.Количество_посещений=" + ABONtextBox4.Text + "";
                    dataAdapterAbonement = new OleDbDataAdapter(selectAbonementWhereCount, connection);
                    dataTableAbonement = new DataTable();
                    dataAdapterAbonement.Fill(dataTableAbonement);
                    ABONmetroGrid1.DataSource = dataTableAbonement;
                    ABONmetroTabControl7.Enabled = false;
                    ABONmetroTabControl6.Enabled = false;
                    ABONmetroTabControl4.Enabled = false;
                    metroTile26.Enabled = false;
                    metroTile5.Enabled = false;
                    if (ABONmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких абонементов не найдено", "Абонемента не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdAbonement();
                    }

                    ABONtextBox4.Text = "";
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile51_Click(object sender, EventArgs e)
        {
            var objSaleAdd = new ModSale {Text = @"Добавить продажу", metroTile1 = {Text = @"Добавить"}};
            connection.Open();
            var cmd = new OleDbCommand("SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия FROM Сотрудник;", connection);
            objSaleAdd.metroComboBox1.DisplayMember = "Фамилия";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objSaleAdd.metroComboBox1.DataSource = list.ToList();
            objSaleAdd.metroComboBox1.DisplayMember = "Value";
            objSaleAdd.metroComboBox1.ValueMember = "Key";
            connection.Close();
            connection.Open();
            var cmd1 =
                new OleDbCommand("SELECT Спортсмен.Идспортсмен, Спортсмен.Фамилия FROM Спортсмен;", connection);
            objSaleAdd.metroComboBox2.DisplayMember = "фамилия";
            var reader1 = cmd1.ExecuteReader();
            var list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int) reader1[0], (string) reader1[1]);
            }

            reader1.Close();
            cmd1.ExecuteNonQuery();
            objSaleAdd.metroComboBox2.DataSource = list1.ToList();
            objSaleAdd.metroComboBox2.DisplayMember = "Value";
            objSaleAdd.metroComboBox2.ValueMember = "Key";
            connection.Close();
            connection.Open();
            var cmd2 =
                new OleDbCommand("SELECT Абонемент.Идабонемент, Абонемент.Название FROM Абонемент;", connection);
            objSaleAdd.metroComboBox3.DisplayMember = "Название";
            var reader2 = cmd2.ExecuteReader();
            var list2 = new Dictionary<int, string>();
            while (reader2.Read())
            {
                list2.Add((int) reader2[0], (string) reader2[1]);
            }

            reader2.Close();
            cmd2.ExecuteNonQuery();
            objSaleAdd.metroComboBox3.DataSource = list2.ToList();
            objSaleAdd.metroComboBox3.DisplayMember = "Value";
            objSaleAdd.metroComboBox3.ValueMember = "Key";
            connection.Close();
            if (objSaleAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    SALEmetroGrid1.Sort(SALEmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryAddSaleAbonement = new OleDbCommand(@"INSERT INTO [Продажа_абонемента]
                                                        ( Идсотрудник, Идспортсмен, Идабонемент, Дата_начала,Дата_окончания)
                                                   VALUES(@idemployee,  @idsportsmen,     @idabonement,       @datezero,        @dateend)",
                        connection);
                    queryAddSaleAbonement.Parameters.AddWithValue("idemployee",
                        Convert.ToInt32(objSaleAdd.metroComboBox1.SelectedValue.ToString()));
                    queryAddSaleAbonement.Parameters.AddWithValue("idsportsmen",
                        Convert.ToInt32(objSaleAdd.metroComboBox2.SelectedValue.ToString()));
                    queryAddSaleAbonement.Parameters.AddWithValue("idabonement",
                        Convert.ToInt32(objSaleAdd.metroComboBox3.SelectedValue.ToString()));
                    queryAddSaleAbonement.Parameters.AddWithValue("datezero",
                        Convert.ToDateTime(objSaleAdd.metroDateTime1.Text));
                    queryAddSaleAbonement.Parameters.AddWithValue("dateend",
                        Convert.ToDateTime(objSaleAdd.metroDateTime2.Text));
                    queryAddSaleAbonement.ExecuteNonQuery();
                    connection.Close();
                    UpdSale();
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile50_Click(object sender, EventArgs e)
        {
            var objSaleUpdate = new ModSale();
            connection.Close();
            objSaleUpdate.Text = @"Редактировать продажи";
            objSaleUpdate.metroTile1.Text = @"Редактировать";
            Debug.Assert(SALEmetroGrid1.CurrentRow != null, "Таблица пуста");
            idSale = Convert.ToInt32(SALEmetroGrid1.CurrentRow.Cells[0].Value);
            connection.Open();
            var cmd = new OleDbCommand("SELECT Сотрудник.Идсотрудник, Сотрудник.Фамилия FROM Сотрудник;", connection);
            objSaleUpdate.metroComboBox1.DisplayMember = "Фамилия";
            var reader = cmd.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }

            reader.Close();
            cmd.ExecuteNonQuery();
            objSaleUpdate.metroComboBox1.DataSource = list.ToList();
            objSaleUpdate.metroComboBox1.DisplayMember = "Value";
            objSaleUpdate.metroComboBox1.ValueMember = "Key";
            connection.Close();
            objSaleUpdate.metroComboBox1.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[7].Value;
            connection.Open();
            var cmd1 =
                new OleDbCommand("SELECT Спортсмен.Идспортсмен, Спортсмен.Фамилия FROM Спортсмен;", connection);
            objSaleUpdate.metroComboBox2.DisplayMember = "фамилия";
            var reader1 = cmd1.ExecuteReader();
            var list1 = new Dictionary<int, string>();
            while (reader1.Read())
            {
                list1.Add((int) reader1[0], (string) reader1[1]);
            }

            reader1.Close();
            cmd1.ExecuteNonQuery();
            objSaleUpdate.metroComboBox2.DataSource = list1.ToList();
            objSaleUpdate.metroComboBox2.DisplayMember = "Value";
            objSaleUpdate.metroComboBox2.ValueMember = "Key";
            connection.Close();
            objSaleUpdate.metroComboBox2.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[8].Value;
            connection.Open();
            var cmd2 =
                new OleDbCommand("SELECT Абонемент.Идабонемент, Абонемент.Название FROM Абонемент;", connection);
            objSaleUpdate.metroComboBox3.DisplayMember = "Название";
            var reader2 = cmd2.ExecuteReader();
            var list2 = new Dictionary<int, string>();
            while (reader2.Read())
            {
                list2.Add((int) reader2[0], (string) reader2[1]);
            }

            reader2.Close();
            cmd2.ExecuteNonQuery();
            objSaleUpdate.metroComboBox3.DataSource = list2.ToList();
            objSaleUpdate.metroComboBox3.DisplayMember = "Value";
            objSaleUpdate.metroComboBox3.ValueMember = "Key";
            connection.Close();
            objSaleUpdate.metroComboBox3.SelectedValue = SALEmetroGrid1.CurrentRow.Cells[9].Value;
            objSaleUpdate.metroDateTime1.Text = Convert.ToString(SALEmetroGrid1.CurrentRow.Cells[4].Value);
            objSaleUpdate.metroDateTime2.Text = Convert.ToString(SALEmetroGrid1.CurrentRow.Cells[5].Value);
            if (objSaleUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    SALEmetroGrid1.Sort(SALEmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateSaleAbonement =
                        new OleDbCommand(
                            "update Продажа_абонемента set Идсотрудник=@idemployee, Идспортсмен=@idsportsmen, Идабонемент=@idabonement, Дата_начала=@datezero,Дата_окончания=@dateend where идпродажа=" +
                            idSale + "", connection);
                    queryUpdateSaleAbonement.Parameters.AddWithValue("idemployee",
                        Convert.ToInt32(objSaleUpdate.metroComboBox1.SelectedValue));
                    queryUpdateSaleAbonement.Parameters.AddWithValue("idsportsmen",
                        Convert.ToInt32(objSaleUpdate.metroComboBox2.SelectedValue));
                    queryUpdateSaleAbonement.Parameters.AddWithValue("idabonement",
                        Convert.ToInt32(objSaleUpdate.metroComboBox3.SelectedValue));
                    queryUpdateSaleAbonement.Parameters.AddWithValue("datezero",
                        Convert.ToDateTime(objSaleUpdate.metroDateTime1.Text));
                    queryUpdateSaleAbonement.Parameters.AddWithValue("dateend",
                        Convert.ToDateTime(objSaleUpdate.metroDateTime2.Text));
                    queryUpdateSaleAbonement.ExecuteNonQuery();
                    connection.Close();
                    UpdSale();
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
        }

        private void metroTile49_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены, что хотите Удалить?",
                    "Подтверждение Удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    connection.Close();
                    connection.Open();
                    Debug.Assert(SALEmetroGrid1.CurrentRow != null, "Таблица пуста");
                    idSale = Convert.ToInt32(SALEmetroGrid1.CurrentRow.Cells[0].Value);
                    var queryDeleteSaleAbonement = new OleDbCommand(@"DELETE FROM продажа_абонемента 
                                                    WHERE идпродажа=" + idSale + "", connection);
                    queryDeleteSaleAbonement.ExecuteNonQuery();
                    UpdSale();
                    connection.Close();
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile47_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.UpdSale();
            string path = Directory.GetCurrentDirectory() + @"\" + "report/SALE.docx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook;
            Worksheet ExcelWorkSheet;
            ExcelWorkBook = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Worksheet) ExcelWorkBook.Worksheets.get_Item(1);
            ExcelWorkSheet.StandardWidth = 17;
            ExcelWorkSheet.Name = "Продажи";
            excelApp.Cells[1, 1] = "Сотрудник";
            excelApp.Cells[1, 2] = "Абонемент";
            excelApp.Cells[1, 3] = "Количество_посещений";
            excelApp.Cells[1, 4] = "Дата_начала";
            excelApp.Cells[1, 5] = "Дата_окончания";
            excelApp.Cells[1, 6] = "Фамилия";
            {
                for (int i = 1; i < SALEmetroGrid1.Rows.Count + 1; i++)
                {
                    excelApp.Cells[i + 1, 1] = SALEmetroGrid1.Rows[i - 1].Cells[1].Value;
                    excelApp.Cells[i + 1, 2] = SALEmetroGrid1.Rows[i - 1].Cells[2].Value;
                    excelApp.Cells[i + 1, 3] = SALEmetroGrid1.Rows[i - 1].Cells[3].Value;
                    excelApp.Cells[i + 1, 4] = SALEmetroGrid1.Rows[i - 1].Cells[4].Value;
                    excelApp.Cells[i + 1, 5] = SALEmetroGrid1.Rows[i - 1].Cells[5].Value;
                    excelApp.Cells[i + 1, 6] = SALEmetroGrid1.Rows[i - 1].Cells[6].Value;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ExcelWorkBook.SaveAs(path);
                    MetroMessageBox.Show(this, "\nОтчет Word сформирование и сохранен по пути" + path,
                        "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (DialogResult.Yes == MetroMessageBox.Show(this,
                        "\nХотите отправить отчет себе на Email?", "Подтверждение отправки", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information))
                {
                    excelApp.Quit();
                    var objEmailReport = new EmailRePort
                    {
                        Text = @"Отправить отчет по продажам", metroTile2 = {Text = @"Отправить"}
                    };
                    if (objEmailReport.ShowDialog() == DialogResult.OK)
                        try
                        {
                            var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                            var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                            var objMessage = new MailMessage(myAddress, toAddress)
                            {
                                Subject = "Отчет по продажам", Body = "Отчет по продажам", IsBodyHtml = true
                            };
                            objMessage.Attachments.Add(new Attachment(path));
                            var smtp = new SmtpClient("smtp.mail.ru", 587)
                            {
                                Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                EnableSsl = true
                            };
                            smtp.Send(objMessage);
                            MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                            objEmailReport.metroTextBox4.Text = "";
                        }
                        catch (Exception exception)
                        {
                            MetroMessageBox.Show(this,
                                @"Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                "Не отправлено",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            FocusMe();
                            if (File.Exists(_fileNameLog) != true)
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                {
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                            else
                            {
                                using (var streamWriter =
                                    new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                {
                                    (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                    streamWriter.WriteLine(_dateLog);
                                    streamWriter.WriteLine(exception.Message);
                                    FocusMe();
                                }
                            }
                        }
                }
            }
        }

        private void metroTile46_Click(object sender, EventArgs e)
        {
            MetroMessageBox.Show(this, "\nОжидайте отчет формируется", "Формирование отчета",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            string path = Directory.GetCurrentDirectory() + @"\" + "report/SALE.docx";
            var wordapp = new Microsoft.Office.Interop.Word.Application();
            wordapp.Visible = true;
            var doc = wordapp.Documents.Add(Visible: true);
            var range = doc.Range();
            try
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(range, SALEmetroGrid1.RowCount + 1, 6);
                table.Borders.Enable = 1;
                table.Cell(1, 1).Range.Text = "Сотрудник";
                table.Cell(1, 2).Range.Text = "Абонемент";
                table.Cell(1, 3).Range.Text = "Количество_посещений";
                table.Cell(1, 4).Range.Text = "Дата_начала";
                table.Cell(1, 5).Range.Text = "Дата_окончания";
                table.Cell(1, 6).Range.Text = "Фамилия";
                table.Range.Bold = 1;
                table.Range.Font.Name = "TimesNewRoman";
                table.Range.Font.Size = 7;
                table.Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        doc.SaveAs(path);
                        MetroMessageBox.Show(this,
                            "\nОтчет Word сформирование и сохранен по пути" + path, "Сохранение", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (DialogResult.Yes == MetroMessageBox.Show(this,
                            "\nХотите отправить отчет себе на Email?", "Подтверждение отправки",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information))
                    {
                        wordapp.Quit();
                        var objEmailReport = new EmailRePort
                        {
                            Text = @"Отправить отчет по продажам", metroTile2 = {Text = @"Отправить"}
                        };
                        if (objEmailReport.ShowDialog() == DialogResult.OK)
                            try
                            {
                                var myAddress = new MailAddress("tempmailgym@mail.ru", "ISgym");
                                var toAddress = new MailAddress(objEmailReport.metroTextBox4.Text);
                                var objMessage = new MailMessage(myAddress, toAddress);
                                objMessage.Subject = "Отчет по продажам";
                                objMessage.Body = "Отчет по продажам";
                                objMessage.IsBodyHtml = true;
                                objMessage.Attachments.Add(new Attachment(path));
                                var smtp = new SmtpClient("smtp.mail.ru", 587)
                                {
                                    Credentials = new NetworkCredential("tempmailgym@mail.ru", "4321Aq"),
                                    EnableSsl = true
                                };
                                smtp.Send(objMessage);
                                MessageBox.Show(@"Сообщение отправлено на почту " + objEmailReport.metroTextBox4.Text);
                                objEmailReport.metroTextBox4.Text = "";
                            }
                            catch (Exception exception)
                            {
                                MetroMessageBox.Show(this,
                                    "Сообщение не отправлено на почту " + objEmailReport.metroTextBox4.Text,
                                    "Не отправлено",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                FocusMe();
                                if (File.Exists(_fileNameLog) != true)
                                {
                                    using (var streamWriter =
                                        new StreamWriter(
                                            new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                                    {
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                                else
                                {
                                    using (var streamWriter =
                                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                                    {
                                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                                        streamWriter.WriteLine(_dateLog);
                                        streamWriter.WriteLine(exception.Message);
                                        FocusMe();
                                    }
                                }
                            }
                    }
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void textBox6_KeyPress_2(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox3_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void metroTextBox2_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= '0' && blockCifr <= '9'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox5_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox4_KeyPress_3(object sender, KeyPressEventArgs e)
        {
            char blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)

                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, Message, Title, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void textBox6_TextChanged_2(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox5_TextChanged_3(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox4_TextChanged_3(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void textBox6_KeyUp_1(object sender, KeyEventArgs e)
        {
            try
            {
                string selectAbonementWhereName =
                    @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник 
                       WHERE[Абонемент.Название] LIKE '%" + SALEtextBox6.Text + "%'";
                dataAdapterSale = new OleDbDataAdapter(selectAbonementWhereName, connection);
                dataTableSale = new DataTable();
                dataAdapterSale.Fill(dataTableSale);
                SALEmetroGrid1.DataSource = dataTableSale;
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl5.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                metroTile42.Enabled = false;
                metroTile43.Enabled = false;
                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапись не найдена", "Абонемента не найдено",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    SALEtextBox6.Text = "";
                    UpdSale();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton5_Click_2(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                string date1 = SALEmetroDateTime2.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                string date2 = SALEmetroDateTime1.Value.ToString("MM/dd/yyyy").Replace('.', '/');
                if (Convert.ToDateTime(SALEmetroDateTime2.Text) < Convert.ToDateTime(SALEmetroDateTime1.Text))
                {
                    SALEmetroTabControl4.Enabled = false;
                    SALEmetroTabControl5.Enabled = false;
                    SALEmetroTabControl7.Enabled = false;
                    metroTile42.Enabled = false;
                    metroTile44.Enabled = false;
                    var adapter = new OleDbDataAdapter(
                        $@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
WHERE Продажа_абонемента.Дата_начала Between #{date1}# and #{date2}#", selectConnection);
                    adapter.Fill(dataSet);
                    SALEmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    if (SALEmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Даты не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSale();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальный период не может быть больше конечного",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdEmployee();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroButton4_Click_3(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                if (Convert.ToInt32(SALEmetroTextBox3.Text) < Convert.ToInt32(SALEmetroTextBox2.Text))
                {
                    var adapter = new OleDbDataAdapter(
                        $@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник where Абонемент.Количество_посещений between {SALEmetroTextBox3.Text} and {SALEmetroTextBox2.Text}",
                        selectConnection);
                    adapter.Fill(dataSet);
                    SALEmetroGrid1.DataSource = dataSet.Tables[0];
                    selectConnection.Close();
                    SALEmetroTabControl4.Enabled = false;
                    SALEmetroTabControl5.Enabled = false;
                    SALEmetroTabControl7.Enabled = false;
                    metroTile43.Enabled = false;
                    metroTile44.Enabled = false;
                    if (SALEmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nЗапись не найдена", "Количество не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSale();
                    }
                }
                else
                {
                    MetroMessageBox.Show(this, "\nНачальное количество не может быть больше конечного",
                        "Ошибка диапазона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdSale();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }


        private void metroButton3_Click_4(object sender, EventArgs e)
        {
            if (SALEtextBox5.Text == "")
            {
                MetroMessageBox.Show(this, @"Пустые(ое) поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SALEmetroTabControl4.Enabled = false;
                    SALEmetroTabControl6.Enabled = false;
                    SALEmetroTabControl7.Enabled = false;
                    metroTile40.Enabled = false;
                    string selectWhereSportsmenSurname =
                        @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
                           WHERE Спортсмен.Фамилия='" + SALEtextBox5.Text + "'";
                    dataAdapterSale = new OleDbDataAdapter(selectWhereSportsmenSurname, connection);
                    dataTableSale = new DataTable();
                    dataAdapterSale.Fill(dataTableSale);
                    SALEmetroGrid1.DataSource = dataTableSale;

                    if (SALEmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSale();
                    }

                    SALEtextBox5.Text = "";
                }

                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroButton2_Click_4(object sender, EventArgs e)
        {
            if (SALEtextBox4.Text == "")
            {
                MetroMessageBox.Show(this, @"Пустые(ое) поля", TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    SALEmetroTabControl4.Enabled = false;
                    SALEmetroTabControl6.Enabled = false;
                    SALEmetroTabControl7.Enabled = false;
                    metroTile41.Enabled = false;
                    string selectWhereEmployeeSurname =
                        @"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия, Абонемент.Название, Абонемент.Количество_посещений, Продажа_абонемента.Дата_начала, Продажа_абонемента.Дата_окончания, Сотрудник.Фамилия, Продажа_абонемента.Идсотрудник, Продажа_абонемента.Идспортсмен, Продажа_абонемента.Идабонемент
FROM Сотрудник INNER JOIN (Абонемент INNER JOIN
(Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен)
ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник
                           WHERE Сотрудник.Фамилия='" + SALEtextBox4.Text + "'";
                    dataAdapterSale = new OleDbDataAdapter(selectWhereEmployeeSurname, connection);
                    dataTableSale = new DataTable();
                    dataAdapterSale.Fill(dataTableSale);
                    SALEmetroGrid1.DataSource = dataTableSale;
                    if (SALEmetroGrid1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "\nТаких спортсменов не найдено", "Спортсмена не найдено",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdSale();
                    }

                    SALEtextBox4.Text = "";
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    if (File.Exists(_fileNameLog) != true)
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                        {
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                    else
                    {
                        using (var streamWriter =
                            new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                        {
                            (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                            streamWriter.WriteLine(_dateLog);
                            streamWriter.WriteLine(exception.Message);
                        }
                    }
                }
            }
        }

        private void metroTile38_Click_2(object sender, EventArgs e)
        {
            try
            {
                var selectConnection = new OleDbConnection(conString);
                selectConnection.Open();
                var dataSet = new DataSet();
                var adapter = new OleDbDataAdapter(
                    @"SELECT Абонемент.Название AS Абонемент, Сотрудник.Фамилия  & Сотрудник.Имя AS [Реализовавший сотрудник], Спортсмен.Фамилия  & Спортсмен.Имя AS Спортсмен, Продажа_абонемента.Дата_начала as Начало,  Продажа_абонемента.Дата_окончания as Конец
                                                                          FROM Спортсмен 
                                                                          INNER JOIN (Сотрудник 
                                                                          INNER JOIN (Абонемент INNER JOIN Продажа_абонемента ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент) 
                                                                          ON Сотрудник.Идсотрудник = Продажа_абонемента.Идсотрудник) 
                                                                          ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;",
                    selectConnection);
                adapter.Fill(dataSet);
                SALEmetroGrid1.DataSource = dataSet.Tables[0];
                selectConnection.Close();
                SALEmetroTabControl4.Enabled = false;
                SALEmetroTabControl5.Enabled = false;
                SALEmetroTabControl6.Enabled = false;
                SALEmetroTabControl9.Enabled = false;
                SALEmetroTabControl7.Enabled = false;
                if (SALEmetroGrid1.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "\nЗапрос не дал результатов", "Запрос пуст",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdSale();
                }
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void сброситьФильтрToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            try
            {
                UpdSale();
            }

            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
                else
                {
                    using (var streamWriter =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (streamWriter.BaseStream).Seek(0, SeekOrigin.End);
                        streamWriter.WriteLine(_dateLog);
                        streamWriter.WriteLine(exception.Message);
                    }
                }
            }
        }

        private void metroTile45_Click(object sender, EventArgs e)
        {
            var reportSportsmen = new ReportSportsmen();
            reportSportsmen.Show();
        }

        private void metroTile28_Click_3(object sender, EventArgs e)
        {
            var recordsOfVisits = new RecordsOfVisits();
            recordsOfVisits.ShowDialog();
        }

        private void HeadForm_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void HeadForm_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/HeadForm.chm"))
                {
                    Help.ShowHelp(null, "Help/HeadForm.chm");
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
                else
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (sw.BaseStream).Seek(0, SeekOrigin.End);
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
            }
        }

        private void HeadForm_Activated(object sender, EventArgs e)
        {
            //SPORTMmetroComboBox1.SelectedIndex = 0;
            //FocusMe();
            try
            {
            //    UpdTrener();
            //    UpdEmployee();
            //    UpdSportsmen();
            //    UpdTrening();
            //    UpdAbonement();
            //    UpdSale();
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (File.Exists(_fileNameLog) != true)
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Create, FileAccess.Write)))
                    {
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
                else
                {
                    using (var sw =
                        new StreamWriter(new FileStream(_fileNameLog, FileMode.Open, FileAccess.Write)))
                    {
                        (sw.BaseStream).Seek(0, SeekOrigin.End);
                        sw.WriteLine(_dateLog);
                        sw.WriteLine(exception.Message);
                        FocusMe();
                    }
                }
            }
        }
    }
}
