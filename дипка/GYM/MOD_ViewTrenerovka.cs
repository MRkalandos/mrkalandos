﻿using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace GYM
{
    public partial class ModViewTrenerovka : MetroFramework.Forms.MetroForm
    {
        private readonly OleDbConnection _connection = new OleDbConnection(
            @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
            Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
            "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        private const string TitleException = "Ошибка";
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Mod_viewTrenerovka.txt";

        public ModViewTrenerovka()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_View_trenerovka_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this,
                    "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            var blockCifr = e.KeyChar;
            if (!(blockCifr >= 'А' && blockCifr <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroFramework.MetroMessageBox.Show(this, "\nНеверный тип данных",
                        TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox1.Text == ""))
                {
                    MetroFramework.MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    _connection.Open();
                    var queryFindClone = new OleDbCommand(@"select *  
                                                                      from [Вид_тренировки] 
                                                                      where название=@name 
                                                                      and идвидтренировка <> " +
                                                          Convert.ToInt32(metroLabel1.Text) + "", _connection);
                    queryFindClone.Parameters.AddWithValue("name", textBox1.Text);
                    queryFindClone.ExecuteNonQuery();
                    if (queryFindClone.ExecuteScalar() != null)
                    {
                        _connection.Close();
                        MetroFramework.MetroMessageBox.Show(this, "\nТакой вид тренировки уже существует",
                            TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        this.DialogResult = DialogResult.OK;
                    }
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Mod_viewTrenerovka.chm"))
                {
                    Help.ShowHelp(null, "Help/Mod_viewTrenerovka.chm");
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Файл не найден", TitleException,MessageBoxButtons.OK,MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void ModViewTrenerovka_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5:
                    button1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    metroTile2.PerformClick();
                    break;
            }
        }

        private void ModViewTrenerovka_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModViewTrenerovka_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
