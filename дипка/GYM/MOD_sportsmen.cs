using System;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModSportsmen : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";
        Inputaccuracy _inputaccuracy = new Inputaccuracy();

        public ModSportsmen()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_Trener_Load(object sender, EventArgs e)
        {
            metroComboBox1.SelectedItem = 0;
            FocusMe();
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

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text == "") ||
                (textBox2.Text == "") ||
                (textBox3.Text == "") ||
                (maskedTextBox1.Text == ""))
            {
                MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModSportsmenInputaccuracyLetter(sender, e);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModSportsmenInputaccuracyLetter(sender, e);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            _inputaccuracy.ModSportsmenInputaccuracyLetter(sender, e);
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

        private void ModSportsmen_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
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

        private void ModSportsmen_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModSportsmen_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}

