using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModAbonement : MetroFramework.Forms.MetroForm
    {
        Inputaccuracy _inputaccuracy = new Inputaccuracy();
        private const string TitleException = "Ошибка";
        private const string Message = @"Неверный тип данных";
        private const string Title = @"Корректность ввода";

        public ModAbonement()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_Abonement_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                (metroComboBox1.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                     Directory.GetParent(Directory.GetCurrentDirectory()).Parent
                                                         ?.FullName + "/ISgym.mdb;Jet OLEDB:Database Password=316206");
                connection.Open();
                var queryFindCloneAbonement = new OleDbCommand(@"select *  
                                                                      from [Абонемент] 
                                                                      where Название=@name
                                                                      and идабонемент <> " +
                                                               Convert.ToInt32(metroLabel4.Text) + "", connection);
                queryFindCloneAbonement.Parameters.AddWithValue("name", textBox1.Text);
                queryFindCloneAbonement.ExecuteNonQuery();
                if (queryFindCloneAbonement.ExecuteScalar() != null)
                {
                    connection.Close();
                    MetroMessageBox.Show(this, "\nТакое название уже существует", TitleException,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this,
                    "\nВы уверены что хотите выйти без сохранения", "Выход", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning))
            {
                Close();
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModAbonementInputaccuracyNumeral(sender, e);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModAbonementInputaccuracyNumeral(sender, e);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private static void NewForm()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};
            form.metroTabControl1.SelectedTab = form.tabPage3;
            ((Control) form.EMPLtabPage6).Enabled = false;
            ((Control) form.tabPage5).Enabled = false;
            ((Control) form.tabPage4).Enabled = false;
            ((Control) form.tabPage29).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.tabPage30).Enabled = false;
            ((Control) form.tabPage27).Enabled = false;
            ((Control) form.tabPage26).Enabled = false;
            ((Control) form.tabPage25).Enabled = false;
            ((Control) form.tabPage24).Enabled = false;
            ((Control) form.tabPage22).Enabled = false;
            form.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm).Start();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("Help/Help.chm"))
                {
                    FocusMe();
                    Help.ShowHelp(null, "Help/Help.chm");
                    FocusMe();
                }
                else
                {
                    MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    FocusMe();
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
        }

        private void ModAbonement_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModAbonement_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModAbonement_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F6:
                    metroButton2.PerformClick();
                    break;
                case Keys.F5:
                    metroTile1.PerformClick();
                    break;
                case Keys.F1:
                    pictureBox1_Click(this, e);
                    break;
                case Keys.Escape:
                    metroTile2.PerformClick();
                    break;
            }
        }
    }
}
