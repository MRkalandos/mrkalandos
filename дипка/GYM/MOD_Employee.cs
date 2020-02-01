using System;
using System.Windows.Forms;
using MetroFramework.Forms;
using System.Data.OleDb;
using System.IO;
using MetroFramework;

namespace GYM
{
    public partial class ModEmployee : MetroForm
    {
        private const string TitleException = "Ошибка";
        Inputaccuracy _inputaccuracy = new Inputaccuracy();

        public ModEmployee()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_Employee_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                (maskedTextBox1.Text == "") ||
                (metroTextBox4.Text == "") ||
                (metroTextBox5.Text == "") ||
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
                var queryFindCloneEmployee = new OleDbCommand(@"select *   from [сотрудник]    
                                                                   where пароль=@pass and
                                                                    идсотрудник <> " + Convert.ToInt32(label1.Text) +
                                                              "", connection);
                queryFindCloneEmployee.Parameters.AddWithValue("pass", metroTextBox5.Text);
                queryFindCloneEmployee.ExecuteNonQuery();
                if (queryFindCloneEmployee.ExecuteScalar() != null)
                {
                    connection.Close();
                    MetroMessageBox.Show(this, "\nТакой пароль уже существует", TitleException,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            var filedialog = new OpenFileDialog();
            filedialog.Filter = @"JPG (*.jpg)|*.jpg|PNG (*.png)|*.png";
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                var fileName = filedialog.SafeFileName;
                var sourcePath = filedialog.FileName;
                var targetPath = @"PhotoEmployee";
                if (!Directory.Exists(targetPath))
                    Directory.CreateDirectory(targetPath);
                var destFile = Path.Combine(targetPath, fileName);
                try
                {
                    File.Copy(sourcePath, destFile, true);
                    metroTextBox6.Text = Path.GetFileName(fileName);
                }
                catch (Exception exception)
                {
                    MetroMessageBox.Show(this, "\nПапка задействована, фото может быть не установлено",
                        TitleException, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    HelperLog.Write(exception.Message);
                    metroTextBox6.Text = Path.GetFileName(fileName);
                }
            }
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

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModEmployeeInputaccuracyLetter(sender, e);
        }

        private void metroTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModEmployeeInputaccuracyLetter(sender, e);
        }

        private void metroTextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModEmployeeInputaccuracyLetter(sender, e);
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModEmployeeInputaccuracyNumeral(sender, e);
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModEmployeeInputaccuracyNumeral(sender, e);
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void metroTextBox3_Click(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            var money = new Money();
            money.ShowDialog();
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

        private void ModEmployee_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModEmployee_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModEmployee_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F6:
                    metroButton1.PerformClick();
                    break;
                case Keys.F7:
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
