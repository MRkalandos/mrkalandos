using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class Money : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        private int _idMoney = 0;
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
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
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
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Debug.Assert(dataGridView1.CurrentRow != null, "Таблица пуста");
            var objMoneyAdd = new ModMoney
            {
                textBox1 = {Text = ""}, Text = @"Добавить зарплату", button1 = {Text = @"Добавить"},
                metroLabel1 = {Text = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value)}
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
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    HelperLog.Write(exception.Message);
                }
                finally
                {
                    FocusMe();
                    connection.Close();
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
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    HelperLog.Write(exception.Message);
                }
                finally
                {
                    FocusMe();
                    connection.Close();
                }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                try
                {
                    if (dataGridView1.RowCount == 0)
                    {
                        MetroMessageBox.Show(this, "Записей больше нет", "Таблица пуста", MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                    else
                    {
                        connection.Close();
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
                    MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    HelperLog.Write(exception.Message);
                }
                finally
                {
                    FocusMe();
                    connection.Close();
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    Help.ShowHelp(null, "Help/Help.chm");
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
            finally
            {
                FocusMe();
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

        private void Money_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
