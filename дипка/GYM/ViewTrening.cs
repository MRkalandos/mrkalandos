using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace GYM
{
    public partial class ViewTrening : MetroFramework.Forms.MetroForm
    {
        public OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                                Directory.GetParent(Directory.GetCurrentDirectory())
                                                                    .Parent?.FullName +
                                                                "/ISgym.mdb;Jet OLEDB:Database Password=316206");

        public DataTable dataTableViewTrening;
        private const string TitleException = "Ошибка";
        public OleDbDataAdapter dataAdapterViewTrening;
        private readonly string _dateLog = DateTime.Now.ToString("dd MMMM yyyy | HH:mm:ss");
        private readonly string _fileNameLog = Directory.GetCurrentDirectory() + @"\" + "LOG/ViewTrening.txt";

        public ViewTrening()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        public void UpdateViewTrening()
        {
            try
            {
                dataAdapterViewTrening =
                    new OleDbDataAdapter(@"SELECT Вид_тренировки.* FROM Вид_тренировки;", connection);
                dataTableViewTrening = new DataTable();
                dataAdapterViewTrening.Fill(dataTableViewTrening);
                ViewmetroGrid1.DataSource = dataTableViewTrening;
                ViewmetroGrid1.Columns[0].Visible = false;
                ViewmetroGrid1.Sort(ViewmetroGrid1.Columns[1], ListSortDirection.Ascending);
                ViewmetroGrid1.Select();
                ViewmetroGrid1.AllowUserToAddRows = false;
                FocusMe();
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

        private void View_trenerovki_Load(object sender, EventArgs e)
        {
            ViewmetroGrid1.Select();
            ViewmetroGrid1.AllowUserToAddRows = false;
            FocusMe();
            try
            {
                UpdateViewTrening();
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

        private void metroButton3_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Close();
                if (DialogResult.Yes == MetroFramework.MetroMessageBox.Show(this, "\nУдалить запись?", "Удалить?",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    connection.Open();
                    Debug.Assert(ViewmetroGrid1.CurrentRow != null, "Таблица пуста");
                    var idViewTrening = Convert.ToInt32(ViewmetroGrid1.CurrentRow.Cells[0].Value);
                    var queryDeleteViewTtrening = new OleDbCommand(@"DELETE FROM Вид_тренировки 
                                                                         WHERE идвидтренировка=" + idViewTrening + "",
                        connection);
                    queryDeleteViewTtrening.ExecuteNonQuery();
                    ViewmetroGrid1.Sort(ViewmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    UpdateViewTrening();
                    connection.Close();
                }
                else
                {
                    FocusMe();
                }
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

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Debug.Assert(ViewmetroGrid1.CurrentRow != null, "Таблица пуста");
            var objViewTreningAdd = new ModViewTrenerovka
            {
                textBox1 = {Text = ""}, Text = @"Добавить тренировку", button1 = {Text = @"Добавить"},
                metroLabel1= { Text = Convert.ToString(ViewmetroGrid1.CurrentRow.Cells[0].Value)}
        };
            if (objViewTreningAdd.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Open();
                    var queryAddViewTtrening = new OleDbCommand(@"INSERT INTO [Вид_тренировки]
                                                                      ( название)
                                                                       VALUES(@name)", connection);
                    queryAddViewTtrening.Parameters.AddWithValue("name", objViewTreningAdd.textBox1.Text);
                    queryAddViewTtrening.ExecuteNonQuery();
                    connection.Close();
                    UpdateViewTrening();
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

        private void metroButton2_Click(object sender, EventArgs e)
        {
            connection.Close();
            var objViewTreningUpdate = new ModViewTrenerovka
            {
                Text = @"Редактировать тренировку", button1 = {Text = @"Редактировать"}
            };

            Debug.Assert(ViewmetroGrid1.CurrentRow != null, "Таблица пуста");
            var idViewTrening = Convert.ToInt32(ViewmetroGrid1.CurrentRow.Cells[0].Value);
            objViewTreningUpdate.metroLabel1.Text = Convert.ToString(ViewmetroGrid1.CurrentRow.Cells[0].Value);
            objViewTreningUpdate.textBox1.Text = Convert.ToString(ViewmetroGrid1.CurrentRow.Cells[1].Value);
            if (objViewTreningUpdate.ShowDialog() == DialogResult.OK)
                try
                {
                    connection.Close();
                    ViewmetroGrid1.Sort(ViewmetroGrid1.Columns[1], ListSortDirection.Ascending);
                    connection.Open();
                    var queryUpdateViewTrening =
                        new OleDbCommand(
                            "update Вид_тренировки set название=@name where идвидтренировка=" + idViewTrening + "",
                            connection);
                    queryUpdateViewTrening.Parameters.AddWithValue("name", objViewTreningUpdate.textBox1.Text);
                    queryUpdateViewTrening.ExecuteNonQuery();
                    connection.Close();
                    UpdateViewTrening();
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/ViewTrening.chm"))
                {
                    FocusMe();
                    Help.ShowHelp(null, "Help/ViewTrening.chm");
                    FocusMe();
                }
                else
                {
                    MetroFramework.MetroMessageBox.Show(this, "Файл не найден", TitleException);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroFramework.MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void ViewTrening_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    metroButton3.PerformClick();
                    FocusMe();
                    break;
                case Keys.F6:
                    metroButton2.PerformClick();
                    FocusMe();
                    break;
                case Keys.F5:
                    metroButton1.PerformClick();
                    FocusMe();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    FocusMe();
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

        private void ViewTrening_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ViewTrening_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
