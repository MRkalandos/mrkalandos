using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class RecordsOfVisits : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        public int idvisit;
        public DataTable dataTableVisits;

        public OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                                Directory.GetParent(Directory.GetCurrentDirectory())
                                                                    .Parent?.FullName +
                                                                "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public OleDbDataAdapter dataAdapterVisits;

        public RecordsOfVisits()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        public void UpdateVisits()
        {
            try
            {
                dataAdapterVisits = new OleDbDataAdapter(
                    @"SELECT  Учет_посещений.Идпосещений , Спортсмен.Фамилия, Абонемент.Название, Продажа_абонемента.Дата_начала as [Дата начала], Продажа_абонемента.Дата_окончания as [Дата окончания], Учет_посещений.Дата as [Дата посещения],Учет_посещений.идпродажа
                                 FROM Абонемент
                                 INNER JOIN ((Спортсмен INNER JOIN Продажа_абонемента 
                                 ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен) INNER JOIN Учет_посещений 
                                 ON Продажа_абонемента.Идпродажа = Учет_посещений.Идпродажа) 
                                 ON Абонемент.Идабонемент = Продажа_абонемента.Идабонемент;",
                    connection);
                dataTableVisits = new DataTable();
                dataAdapterVisits.Fill(dataTableVisits);
                VISITSGrid.DataSource = dataTableVisits;
                VISITSGrid.Columns[0].Visible = false;
                VISITSGrid.Columns[6].Visible = false;
                VISITSGrid.Sort(VISITSGrid.Columns[1], ListSortDirection.Ascending);
                VISITSGrid.Select();
                VISITSGrid.AllowUserToAddRows = false;
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


        private void VISITS_Load(object sender, EventArgs e)
        {
            try
            {
                UpdateVisits();
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

        private void metroButton3_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (VISITSGrid.RowCount == 0)
                {
                    MetroMessageBox.Show(this, "Записей больше нет", "Таблица пуста", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        connection.Close();
                        connection.Open();
                        Debug.Assert(VISITSGrid.CurrentRow != null, "Таблица пуста");
                        var idVisit = Convert.ToInt32(VISITSGrid.CurrentRow.Cells[0].Value);
                        var queryDeleteVisit = new OleDbCommand(@"DELETE FROM учет_посещений 
                                                    WHERE идпосещений=" + idVisit + "", connection);
                        queryDeleteVisit.ExecuteNonQuery();
                        connection.Close();
                        VISITSGrid.Sort(VISITSGrid.Columns[1], ListSortDirection.Ascending);
                        UpdateVisits();
                    }
                    catch (Exception exception)
                    {
                        MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        HelperLog.Write(exception.Message);
                    }
                    finally
                    {
                        connection.Close();
                        FocusMe();
                    }
                }
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            var objVisiTeAdd = new MOD_VISITS {Text = @"Добавить посещение", metroTile1 = {Text = @"Добавить"}};
            connection.Open();
            var dbCommandVisit = new OleDbCommand(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия
FROM Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;", connection);
            objVisiTeAdd.metroComboBox1.DisplayMember = "Спортсмен.Фамилия";
            var reader = dbCommandVisit.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }
            reader.Close();
            dbCommandVisit.ExecuteNonQuery();
            objVisiTeAdd.metroComboBox1.DataSource = list.ToList();
            objVisiTeAdd.metroComboBox1.DisplayMember = "Value";
            objVisiTeAdd.metroComboBox1.ValueMember = "Key";
            connection.Close();
            if (objVisiTeAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    VISITSGrid.Sort(VISITSGrid.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryInsertVisit = new OleDbCommand(@"INSERT INTO [учет_посещений]
                                                        ( Дата,Идпродажа)
                                                        VALUES(@sdate,@idsale)", connection);
                    queryInsertVisit.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objVisiTeAdd.metroDateTime1.Text));
                    queryInsertVisit.Parameters.AddWithValue("idsale",
                        Convert.ToInt32(objVisiTeAdd.metroComboBox1.SelectedValue.ToString()));
                    queryInsertVisit.ExecuteNonQuery();
                    connection.Close();
                    UpdateVisits();
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

        private void metroButton2_Click(object sender, EventArgs e)
        {
            var objVisiteUpdate = new MOD_VISITS();
            connection.Close();
            objVisiteUpdate.Text = @"Редактировать посещение";
            objVisiteUpdate.metroTile1.Text = @"Редактировать";
            Debug.Assert(VISITSGrid.CurrentRow != null, "Таблица пуста");
            idvisit = Convert.ToInt32(VISITSGrid.CurrentRow.Cells[0].Value);
            objVisiteUpdate.metroDateTime1.Text = Convert.ToString(VISITSGrid.CurrentRow.Cells[5].Value);
            connection.Open();
            var oleDbCommand = new OleDbCommand(@"SELECT Продажа_абонемента.Идпродажа, Спортсмен.Фамилия
FROM Спортсмен INNER JOIN Продажа_абонемента ON Спортсмен.Идспортсмен = Продажа_абонемента.Идспортсмен;", connection);
            objVisiteUpdate.metroComboBox1.DisplayMember = "Спортсмен.Фамилия";
            var reader = oleDbCommand.ExecuteReader();
            var list = new Dictionary<int, string>();
            while (reader.Read())
            {
                list.Add((int) reader[0], (string) reader[1]);
            }
            reader.Close();
            oleDbCommand.ExecuteNonQuery();
            objVisiteUpdate.metroComboBox1.DataSource = list.ToList();
            objVisiteUpdate.metroComboBox1.DisplayMember = "Value";
            objVisiteUpdate.metroComboBox1.ValueMember = "Key";
            connection.Close();
            objVisiteUpdate.metroComboBox1.SelectedValue = VISITSGrid.CurrentRow.Cells[6].Value;
            if (objVisiteUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    VISITSGrid.Sort(VISITSGrid.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateVisit =
                        new OleDbCommand(
                            "update Учет_посещений set дата=@date, идпродажа=@idsale where идпосещений=" + idvisit +
                            "", connection);
                    queryUpdateVisit.Parameters.AddWithValue("date",
                        Convert.ToDateTime(objVisiteUpdate.metroDateTime1.Text));
                    queryUpdateVisit.Parameters.AddWithValue("idsale",
                        Convert.ToInt32(objVisiteUpdate.metroComboBox1.SelectedValue));
                    queryUpdateVisit.ExecuteNonQuery();
                    connection.Close();
                    UpdateVisits();
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

        private void Visits_KeyDown(object sender, KeyEventArgs e)
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

        private void RecordsOfVisits_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void RecordsOfVisits_Click(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void VISITSGrid_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}

