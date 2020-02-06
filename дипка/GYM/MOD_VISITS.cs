using System;
using System.IO;
using System.Windows.Forms;
using MetroFramework;

namespace GYM
{
    public partial class MOD_VISITS : MetroFramework.Forms.MetroForm
    {
        private const string TitleException = "Ошибка";

        public MOD_VISITS()
        {
            InitializeComponent();
            this.KeyPreview = true;
        }

        private void MOD_VISITS_Load(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((metroDateTime1.Text == "") ||
                    (metroComboBox1.Text == ""))
                {
                    MetroMessageBox.Show(this, "\nНе все поля заполнены", TitleException,
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
            finally
            {
                FocusMe();
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
            else
            {
                FocusMe();
            }
        }

        private static void NewForm2()
        {
            var form = new HeadForm {pictureBox2 = {Visible = true}};
            form.metroTabControl1.SelectedTab = form.tabPage2;
            form.metroTabControl2.SelectedTab = form.tabPage30;
            ((Control) form.tabPage1).Enabled = false;
            ((Control) form.tabPage3).Enabled = false;
            ((Control) form.tabPage13).Enabled = false;
            ((Control) form.tabPage29).Enabled = false;
            ((Control) form.tabPage40).Enabled = false;
            ((Control) form.tabPage39).Enabled = false;
            ((Control) form.tabPage38).Enabled = false;
            ((Control) form.tabPage37).Enabled = false;
            ((Control) form.tabPage36).Enabled = false;
            form.ShowInTaskbar = false;
            form.ControlBox = false;
            form.ShowDialog();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(NewForm2).Start();
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

        private void MOD_VISITS_Shown(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void MOD_VISITS_Activated(object sender, EventArgs e)
        {
            FocusMe();
        }

        private void MOD_VISITS_KeyDown(object sender, KeyEventArgs e)
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

        private void MOD_VISITS_Click(object sender, EventArgs e)
        {
            FocusMe();
        }
    }
}
