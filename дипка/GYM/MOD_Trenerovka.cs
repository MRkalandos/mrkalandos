using System;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class ModTrenerovka : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public ModTrenerovka()
        {
            InitializeComponent();
        }

        private void MOD_Trenerovka_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MetroMessageBox.Show(this, "\nВы уверены что хотите выйти без сохранения", "Выход",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                Close();
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox1.Text == "") ||
                    (metroComboBox1.Text == "") ||
                    (metroComboBox2.Text == ""))

                {
                    MetroMessageBox.Show(this, "\nНе все поля заполнены", "Корректность",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                }
            }
            catch (Exception exception)
            {
                MetroMessageBox.Show(this, exception.Message, TitleException, MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                HelperLog.Write(exception.Message);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            var keyChar = e.KeyChar;
            if (!(keyChar >= 'А' && keyChar <= 'я'))
            {
                if (e.KeyChar != (char) Keys.Back)
                {
                    e.Handled = true;
                    MetroMessageBox.Show(this, "Неверный тип данных", TitleException,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox) sender).Text.Length == 1)
                ((TextBox) sender).Text = ((TextBox) sender).Text.ToUpper();
            ((TextBox) sender).Select(((TextBox) sender).Text.Length, 0);
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            var view = new ViewTrening();
            view.Show();
        }

        private static void NewForm()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};
            form.metroTabControl1.SelectedTab = form.tabPage1;
            form.metroTabControl2.SelectedTab = form.tabPage5;
            ((Control) form.tabPage2).Enabled = false;
            ((Control) form.tabPage3).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.EMPLtabPage6).Enabled = false;
            ((Control) form.tabPage4).Enabled = false;
            ((Control) form.tabPage20).Enabled = false;
            ((Control) form.tabPage21).Enabled = false;
            ((Control) form.tabPage18).Enabled = false;
            ((Control) form.tabPage6).Enabled = false;
            form.ShowDialog();
        }

        public void metroButton2_Click(object sender, EventArgs e)
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
                    MetroFramework.MetroMessageBox.Show(this, "Файл не найден", TitleException, MessageBoxButtons.OK,
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

        private void ModTrenerovka_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModTrenerovka_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void ModTrenerovka_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F5:
                    metroTile1.PerformClick();
                    break;
                case Keys.F6:
                    metroButton1.PerformClick();
                    break;
                case Keys.F7:
                    metroButton2.PerformClick();
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
    

