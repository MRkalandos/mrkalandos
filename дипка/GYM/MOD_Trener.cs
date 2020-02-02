using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModTrener : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        readonly Inputaccuracy _inputaccuracy = new Inputaccuracy();

        public ModTrener()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_Trener_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox1.Text == "") ||
                    (textBox2.Text == "") ||
                    (textBox3.Text == "") ||
                    (maskedTextBox1.Text == "") ||
                    (metroTextBox4.Text == "") ||
                    (metroTextBox5.Text == "") ||
                    (metroTextBox1.Text == ""))
                {
                    MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    var connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                                                         Directory.GetParent(Directory.GetCurrentDirectory()).Parent
                                                             ?.FullName +
                                                         "/ISgym.mdb;Jet OLEDB:Database Password=316206");
                    connection.Open();
                    var queryFindClone = new OleDbCommand(@"select *  
                                                                      from [тренер] 
                                                                      where пароль=@pass and
                                                                    идтренер <> " + Convert.ToInt32(label1.Text) + "",
                        connection);
                    queryFindClone.Parameters.AddWithValue("pass", metroTextBox5.Text);
                    queryFindClone.ExecuteNonQuery();
                    if (queryFindClone.ExecuteScalar() != null)
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

        private void metroButton1_Click(object sender, EventArgs e)
        {
            var filedialog = new OpenFileDialog {Filter = @"JPG (*.jpg)|*.jpg|PNG (*.png)|*.png"};
            if (filedialog.ShowDialog() == DialogResult.OK)
            {
                var fileName = filedialog.SafeFileName;
                var sourcePath = filedialog.FileName;
                var targetPath = @"PhotoTrener";
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
                    MetroMessageBox.Show(this, "\nПапка задействована, фото может быть не установлено", TitleException,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    metroTextBox6.Text = Path.GetFileName(fileName);
                    HelperLog.Write(exception.Message);
                }
                finally
                {
                    FocusMe();
                }
            }
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
            else
            {
                FocusMe();
            }
        }


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyLetter(sender, e);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyLetter(sender, e);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyLetter(sender, e);
        }

        private void maskedTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyNumeral(sender, e);
        }

        private void metroTextBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyNumeral(sender, e);
        }

        private void metroTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModTrenerInputaccuracyNumeral(sender, e);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            _inputaccuracy.UpperLetter(sender, e);
        }

        private void MOD_Trener_Shown(object sender, EventArgs e)
        {
            FocusMe();
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

        private void MOD_Trener_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModTrener_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F6:
                    metroButton1.PerformClick();
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

        private void ModTrener_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
