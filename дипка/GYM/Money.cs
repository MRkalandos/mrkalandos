using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace GYM
{
    public partial class Money : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        private int _idMoney = 0;
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/Money.txt";
        public DataTable dataTableMoney;

        public string conString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                   Directory.GetParent(Directory.GetCurrentDirectory()).Parent?.FullName +
                                   "/ISgym.mdb;Jet OLEDB:Database Password=316206");


        public OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                                Directory.GetParent(Directory.GetCurrentDirectory())
                                                                    .Parent?.FullName +
                                                                "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public OleDbDataAdapter dataAdapterMoney;

        public Money()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        public void UpdateMoney()
        {
            try
            {
                dataAdapterMoney = new OleDbDataAdapter(@"SELECT * from зарплата_сотрудника;", connection);
                dataTableMoney = new DataTable();
                dataAdapterMoney.Fill(dataTableMoney);
                dataGridView1.DataSource = dataTableMoney;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                dataGridView1.Select();
                dataGridView1.AllowUserToAddRows = false;
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

        private void Money_Load(object sender, EventArgs e)
        {
            try
            {
                this.FormBorderStyle = FormBorderStyle.None;
                UpdateMoney();
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

        private void button1_Click(object sender, EventArgs e)
        {
            var objMoneyAdd = new ModMoney
            {
                textBox1 = {Text = ""}, Text = @"Добавить зарплату", button1 = {Text = @"Добавить"},
                 metroLabel1= { Text = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value)}
        };
            if (objMoneyAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Open();
                    var queryAddMoney = new OleDbCommand(@"INSERT INTO [Зарплата_сотрудника]
                                                        ( Зарплата)
                                                        VALUES(@money)", connection);
                    queryAddMoney.Parameters.AddWithValue("money", objMoneyAdd.textBox1.Text);
                    queryAddMoney.ExecuteNonQuery();
                    connection.Close();
                    UpdateMoney();
                }
                catch (Exception exception)
                {
                    MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void button2_Click(object sender, EventArgs e)
        {
            connection.Close();
            var objMoneyUpdate = new ModMoney {Text = @"Редактировать зарплату", button1 = {Text = @"Редактировать"}};
            Debug.Assert(dataGridView1.CurrentRow != null, "Таблица пуста");
            _idMoney = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            objMoneyUpdate.metroLabel1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            objMoneyUpdate.textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            if (objMoneyUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateMoney =
                        new OleDbCommand(
                            "update зарплата_сотрудника set зарплата=@money where идзарплата=" + _idMoney + "",
                            connection);
                    queryUpdateMoney.Parameters.AddWithValue("money", objMoneyUpdate.textBox1.Text);
                    queryUpdateMoney.ExecuteNonQuery();
                    connection.Close();
                    UpdateMoney();
                }
                catch (Exception exception)
                {
                    MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Close();
                if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    connection.Open();
                    Debug.Assert(dataGridView1.CurrentRow != null, "Таблица пуста");
                    _idMoney = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                    var queryDeleteMoney = new OleDbCommand(@"DELETE FROM зарплата_сотрудника 
                                                    WHERE идзарплата=" + _idMoney + "", connection);
                    queryDeleteMoney.ExecuteNonQuery();
                    dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
                    UpdateMoney();
                    connection.Close();
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this,
                    "\nУдалить запись нельзя, данная запись используется в таблице 'Cотрудник'",TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Money.chm"))
                {
                    Help.ShowHelp(null, "Help/Money.chm");
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

        private void Money_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void Money_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void Money_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    metroButton3.PerformClick();
                    break;
                case Keys.F6:
                    metroButton2.PerformClick();
                    break;
                case Keys.F5:
                    metroButton1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    Close();
                    break;
            }
        }

        private void metroTabControl1_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void Money_FormClosed(object sender, FormClosedEventArgs e)
        {
        }
    }
}
